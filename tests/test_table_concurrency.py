"""Two 'users' working on the same table through their own Table objects, in one process.

Alice and Bob each read the table, edit their own copy, and write it back. The
second writer's write() must reconcile with what the first already saved.
"""

from __future__ import annotations

import warnings
from datetime import datetime, timezone

import pytest
from openpyxl import load_workbook

from pyhandlexl import (
    ColumnType,
    ColumnTypeError,
    MergeConflictWarning,
    MergeReport,
    Table,
    TableNotFoundError,
    TableStyle,
    create_sheet,
    delete_table,
    table_info,
)


@pytest.fixture
def path(book):
    create_sheet(book, "Data")
    Table(
        data=[[1, 2, 3], [4, 5, 6], [7, 8, 9]],
        row_labels=["r1", "r2", "r3"],
        corner="id",
        column_headers=["a", "b", "c"],
        name="T",
    ).create(book, sheet="Data")
    return book


@pytest.fixture
def users(path):
    """Alice and Bob, each holding their own freshly read copy."""
    return Table.read(path, "T"), Table.read(path, "T")


def written(path):
    return Table.read(path, "T")


def quiet_write(table, path):
    with warnings.catch_warnings():
        warnings.simplefilter("error", MergeConflictWarning)
        table.write(path)


class TestNoOneElseWrote:
    def test_write_reports_no_merge(self, path):
        t = Table.read(path, "T")
        t.set_cell(row="r1", column="a", value=100)
        quiet_write(t, path)
        assert t.last_merge is None
        assert written(path).read_cell(row="r1", column="a") == 100

    def test_a_write_that_only_touches_our_own_earlier_save_is_not_a_merge(self, path):
        t = Table.read(path, "T")
        for i in range(3):
            t.set_cell(row="r1", column="a", value=i)
            quiet_write(t, path)
            assert t.last_merge is None

    def test_last_merge_is_none_before_any_write(self, path):
        assert Table.read(path, "T").last_merge is None


class TestAppends:
    def test_both_append_different_rows(self, path, users):
        alice, bob = users
        alice.add_row("alice", [1, 1, 1])
        bob.add_row("bob", [2, 2, 2])
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.row_labels == ["r1", "r2", "r3", "alice", "bob"]
        assert final.read_row("alice") == [1, 1, 1]
        assert final.read_row("bob") == [2, 2, 2]

    def test_the_second_writers_object_ends_up_holding_the_merged_table(self, path, users):
        alice, bob = users
        alice.add_row("alice", [1, 1, 1])
        bob.add_row("bob", [2, 2, 2])
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert bob.data.row_labels == ["r1", "r2", "r3", "alice", "bob"]
        assert bob.last_merge == MergeReport(rows_added=("alice",))

    def test_after_a_merge_the_next_write_has_a_fresh_base(self, path, users):
        alice, bob = users
        alice.add_row("alice", [1, 1, 1])
        quiet_write(alice, path)
        bob.add_row("bob", [2, 2, 2])
        quiet_write(bob, path)
        bob.add_row("bob2", [3, 3, 3])
        quiet_write(bob, path)
        assert bob.last_merge is None  # nothing new on disk since the last write
        assert written(path).data.row_labels[-2:] == ["bob", "bob2"]

    def test_many_sequential_appenders_from_stale_copies(self, path):
        copies = [Table.read(path, "T") for _ in range(6)]
        for i, t in enumerate(copies):
            t.add_row(f"u{i}", [i, i, i])
            quiet_write(t, path)
        assert written(path).data.row_labels == ["r1", "r2", "r3", *[f"u{i}" for i in range(6)]]

    def test_both_insert_at_the_top(self, path, users):
        alice, bob = users
        alice.insert_row(1, "alice", [1, 1, 1])
        bob.insert_row(1, "bob", [2, 2, 2])
        quiet_write(alice, path)
        quiet_write(bob, path)
        labels = written(path).data.row_labels
        assert labels[0] == "bob" and set(labels) == {"alice", "bob", "r1", "r2", "r3"}

    def test_our_row_inserted_in_the_middle_stays_after_its_predecessor(self, path, users):
        alice, bob = users
        alice.add_row("alice", [1, 1, 1])
        bob.insert_row(2, "bob", [2, 2, 2])  # between r1 and r2
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["r1", "bob", "r2", "r3", "alice"]

    def test_both_add_the_same_label_with_the_same_values_is_fine(self, path, users):
        alice, bob = users
        for t in (alice, bob):
            t.add_row("same", [0, 0, 0])
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels.count("same") == 1

    def test_both_add_the_same_label_with_different_values_ours_wins_with_a_warning(
        self, path, users
    ):
        alice, bob = users
        alice.add_row("same", [1, 1, 1])
        bob.add_row("same", [2, 2, 2])
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning) as caught:
            bob.write(path)
        assert written(path).read_row("same") == [2, 2, 2]
        assert len(bob.last_merge.conflicts) == 3
        assert "'same'" in str(caught[0].message)


class TestCellEdits:
    def test_edits_to_different_cells_both_survive(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value="alice")
        bob.set_cell(row="r3", column="c", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.read_cell(row="r1", column="a") == "alice"
        assert final.read_cell(row="r3", column="c") == "bob"

    def test_bobs_untouched_cells_take_alices_values(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value="alice")
        bob.set_cell(row="r3", column="c", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert bob.read_cell(row="r1", column="a") == "alice"
        assert bob.last_merge.cells_updated == 1

    def test_editing_the_same_cell_ours_wins_and_names_the_cell(self, path, users):
        alice, bob = users
        alice.set_cell(row="r2", column="b", value="alice")
        bob.set_cell(row="r2", column="b", value="bob")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match=r"cell \('r2', 'b'\)"):
            bob.write(path)
        assert written(path).read_cell(row="r2", column="b") == "bob"

    def test_same_value_in_the_same_cell_is_no_conflict(self, path, users):
        alice, bob = users
        alice.set_cell(row="r2", column="b", value=42)
        bob.set_cell(row="r2", column="b", value=42)
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).read_cell(row="r2", column="b") == 42

    def test_a_whole_row_replacement_merges_cell_by_cell(self, path, users):
        alice, bob = users
        alice.set_row("r1", ["A", "A", "A"])
        bob.set_cell(row="r1", column="c", value="B")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning):
            bob.write(path)
        assert written(path).read_row("r1") == ["A", "A", "B"]

    def test_typed_values_survive_the_merge(self, path, users):
        alice, bob = users
        when = datetime(2026, 5, 17, 8, 30, 15)
        alice.set_cell(row="r1", column="a", value=when)
        alice.set_cell(row="r1", column="b", value=True)
        alice.set_cell(row="r1", column="c", value=2.5)
        bob.set_cell(row="r2", column="a", value="x")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.read_row("r1") == [when, True, 2.5]
        assert final.read_cell(row="r1", column="b") is True

    def test_bool_versus_int_is_seen_as_a_real_difference(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value=True)  # 1 -> True
        bob.set_cell(row="r1", column="a", value=7)
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning):
            bob.write(path)
        assert written(path).read_cell(row="r1", column="a") == 7

    def test_excel_float_round_trip_does_not_look_like_an_edit(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value=1.0)  # 1 -> 1.0 reads back as 1
        bob.set_cell(row="r1", column="a", value=99)
        quiet_write(alice, path)
        quiet_write(bob, path)  # alice's "edit" was a no-op, so no conflict
        assert written(path).read_cell(row="r1", column="a") == 99

    def test_a_datetime_survives_a_merge_without_phantom_conflicts(self, path):
        when = datetime(2026, 1, 2, 3, 4, 5, 123456, tzinfo=None)
        seed = Table.read(path, "T")
        seed.set_cell(row="r1", column="a", value=when)
        quiet_write(seed, path)
        alice, bob = Table.read(path, "T"), Table.read(path, "T")
        alice.set_cell(row="r2", column="a", value="alice")
        bob.set_cell(row="r3", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)  # would warn if the datetime looked edited
        assert isinstance(written(path).read_cell(row="r1", column="a"), datetime)


class TestDeletes:
    def test_they_delete_a_row_we_never_touched(self, path, users):
        alice, bob = users
        alice.drop_row("r2")
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["r1", "r3"]
        assert bob.last_merge.rows_removed == ("r2",)

    def test_we_delete_a_row_they_never_touched(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value="alice")
        bob.drop_row("r2")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["r1", "r3"]

    def test_they_delete_a_row_we_edited_ours_is_restored_with_a_warning(self, path, users):
        alice, bob = users
        alice.drop_row("r2")
        bob.set_cell(row="r2", column="b", value="bob edit")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match="deleted by another writer but edited"):
            bob.write(path)
        final = written(path)
        assert "r2" in final.data.row_labels
        assert final.read_cell(row="r2", column="b") == "bob edit"

    def test_we_delete_a_row_they_edited_ours_wins_with_a_warning(self, path, users):
        alice, bob = users
        alice.set_cell(row="r2", column="b", value="alice edit")
        bob.drop_row("r2")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match="deleted by you but edited"):
            bob.write(path)
        assert "r2" not in written(path).data.row_labels

    def test_both_delete_the_same_row(self, path, users):
        alice, bob = users
        alice.drop_row("r2")
        bob.drop_row("r2")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["r1", "r3"]

    def test_delete_then_re_add_the_same_label(self, path, users):
        alice, bob = users
        bob.drop_row("r2")
        bob.add_row("r2", ["new", "new", "new"])
        alice.set_cell(row="r1", column="a", value="alice")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.read_row("r2") == ["new", "new", "new"]
        assert final.read_cell(row="r1", column="a") == "alice"

    def test_a_table_can_be_emptied_by_one_writer_while_another_appends(self, path, users):
        alice, bob = users
        for label in ("r1", "r2", "r3"):
            alice.drop_row(label)
        bob.add_row("bob", [0, 0, 0])
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["bob"]


class TestRenames:
    def test_our_rename_and_their_edit_of_the_same_row_both_survive(self, path, users):
        alice, bob = users
        alice.set_cell(row="r2", column="b", value="alice edit")
        bob.rename_row("r2", "renamed")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.row_labels == ["r1", "renamed", "r3"]
        assert final.read_cell(row="renamed", column="b") == "alice edit"

    def test_their_rename_arrives_as_a_replaced_row(self, path, users):
        alice, bob = users
        alice.rename_row("r2", "renamed")
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.row_labels == ["r1", "renamed", "r3"]

    def test_renaming_a_row_twice_still_tracks_the_original(self, path, users):
        alice, bob = users
        alice.set_cell(row="r2", column="b", value="alice edit")
        bob.rename_row("r2", "tmp")
        bob.rename_row("tmp", "final")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.row_labels == ["r1", "final", "r3"]
        assert final.read_cell(row="final", column="b") == "alice edit"

    def test_swapping_two_labels_does_not_confuse_the_merge(self, path, users):
        alice, bob = users
        bob.rename_row("r1", "tmp")
        bob.rename_row("r2", "r1")
        bob.rename_row("tmp", "r2")
        alice.set_cell(row="r3", column="a", value="alice")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.row_labels == ["r2", "r1", "r3"]
        assert final.read_cell(row="r1", column="a") == 4  # what was r2's data
        assert final.read_cell(row="r3", column="a") == "alice"

    def test_a_rename_onto_a_label_they_added_ours_wins(self, path, users):
        alice, bob = users
        alice.add_row("clash", [9, 9, 9])
        bob.rename_row("r1", "clash")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match="renamed"):
            bob.write(path)
        final = written(path)
        assert final.data.row_labels.count("clash") == 1
        assert final.read_row("clash") == [1, 2, 3]

    def test_column_rename_keeps_their_edits(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="b", value="alice")
        bob.rename_column("b", "beta")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.column_headers == ["a", "beta", "c"]
        assert final.read_cell(row="r1", column="beta") == "alice"


class TestColumns:
    def test_both_add_different_columns(self, path, users):
        alice, bob = users
        alice.add_column("alice", ["A1", "A2", "A3"])
        bob.add_column("bob", ["B1", "B2", "B3"])
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.column_headers == ["a", "b", "c", "alice", "bob"]
        assert final.read_column("alice") == ["A1", "A2", "A3"]
        assert final.read_column("bob") == ["B1", "B2", "B3"]

    def test_a_column_grows_and_a_row_is_appended_by_the_other_writer(self, path, users):
        alice, bob = users
        alice.add_column("extra", ["x", "y", "z"])
        bob.add_row("bob", [0, 0, 0])
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.read_row("bob") == [0, 0, 0, None]  # bob's row has no value for 'extra'
        assert final.read_column("extra") == ["x", "y", "z", None]

    def test_they_drop_a_column_we_never_touched(self, path, users):
        alice, bob = users
        alice.drop_column("b")
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.column_headers == ["a", "c"]
        assert bob.last_merge.columns_removed == ("b",)

    def test_they_drop_a_column_we_edited_ours_is_restored_with_a_warning(self, path, users):
        alice, bob = users
        alice.drop_column("b")
        bob.set_cell(row="r1", column="b", value="bob")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match="deleted by another writer but edited"):
            bob.write(path)
        assert written(path).data.column_headers == ["a", "b", "c"]
        assert written(path).read_cell(row="r1", column="b") == "bob"

    def test_a_column_inserted_in_the_middle_keeps_its_place(self, path, users):
        alice, bob = users
        alice.add_column("alice", [0, 0, 0])
        bob.insert_column(2, "bob", [1, 1, 1])
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.column_headers == ["a", "bob", "b", "c", "alice"]

    def test_merged_columns_growing_the_table_shift_the_neighbour_table(self, path):
        Table(data=[["n"]], row_labels=["k"], column_headers=["v"], name="N").create(
            path, sheet="Data"
        )
        alice, bob = Table.read(path, "T"), Table.read(path, "T")
        alice.add_column("a2", [0, 0, 0])
        bob.add_column("b2", [0, 0, 0])
        quiet_write(alice, path)
        quiet_write(bob, path)  # now three columns wider than it started
        assert written(path).data.column_headers == ["a", "b", "c", "a2", "b2"]
        neighbour = Table.read(path, "N")
        assert neighbour.read_cell(row="k", column="v") == "n"
        assert table_info(path, "N").n_rows == 1

    def test_a_shrinking_merge_leaves_the_table_consistent(self, path, users):
        alice, bob = users
        alice.drop_column("c")
        bob.drop_column("b")
        quiet_write(alice, path)
        quiet_write(bob, path)
        final = written(path)
        assert final.data.column_headers == ["a"]
        assert final.read_column("a") == [1, 4, 7]
        assert table_info(path, "T").n_cols == 1
        # and nothing stale is left in the freed cells
        ws = load_workbook(path)["Data"]
        assert ws.max_column <= 2


class TestCornerStyleTypes:
    def test_their_corner_change_is_kept_when_we_did_not_touch_it(self, path, users):
        alice, bob = users
        alice.set_corner("alice")
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).data.corner == "alice"

    def test_both_change_the_corner_ours_wins(self, path, users):
        alice, bob = users
        alice.set_corner("alice")
        bob.set_corner("bob")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning, match="corner"):
            bob.write(path)
        assert written(path).data.corner == "bob"

    def test_their_style_change_is_kept_when_we_did_not_touch_it(self, path, users):
        alice, bob = users
        alice.style = TableStyle.MINIMAL
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).style == TableStyle.MINIMAL

    def test_our_style_change_survives(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value="alice")
        bob.style = TableStyle.NONE
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).style == TableStyle.NONE

    def test_their_column_type_is_kept_and_enforced(self, path, users):
        alice, bob = users
        alice.set_column_type("a", ColumnType.NUMBER)
        bob.set_cell(row="r1", column="b", value="bob")
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert written(path).column_types["a"] == ColumnType.NUMBER

    def test_a_restriction_they_added_rejects_our_value_after_the_merge(self, path, users):
        alice, bob = users
        alice.set_column_type("a", ColumnType.NUMBER)
        bob.set_cell(row="r1", column="a", value="not a number")
        quiet_write(alice, path)
        before = (bob.data.rows, bob.column_types)
        with pytest.raises(ColumnTypeError):
            bob.write(path)
        # this object is exactly as the caller left it, and the file is as alice wrote it
        assert (bob.data.rows, bob.column_types) == before
        assert written(path).read_cell(row="r1", column="a") == 1

    def test_a_failed_write_can_be_fixed_and_retried(self, path, users):
        alice, bob = users
        alice.set_column_type("a", ColumnType.NUMBER)
        bob.set_cell(row="r1", column="a", value="not a number")
        quiet_write(alice, path)
        with pytest.raises(ColumnTypeError):
            bob.write(path)
        bob.set_cell(row="r1", column="a", value=5)
        quiet_write(bob, path)
        final = written(path)
        assert final.read_cell(row="r1", column="a") == 5
        assert final.column_types["a"] == ColumnType.NUMBER


class TestMetadata:
    def test_created_at_is_preserved_and_modified_at_moves_on(self, path, users):
        alice, bob = users
        original = table_info(path, "T")
        alice.set_cell(row="r1", column="a", value=1)
        quiet_write(alice, path)
        bob.set_cell(row="r2", column="a", value=1)
        quiet_write(bob, path)
        info = table_info(path, "T")
        assert info.created_at == original.created_at
        assert info.modified_at >= original.modified_at
        assert info.modified_at <= datetime.now(timezone.utc)

    def test_info_counts_match_the_merged_table(self, path, users):
        alice, bob = users
        alice.add_row("alice", [0, 0, 0])
        bob.add_row("bob", [0, 0, 0])
        quiet_write(alice, path)
        quiet_write(bob, path)
        assert table_info(path, "T").n_rows == 5


class TestWhatHasNothingToMergeAgainst:
    def test_a_table_never_read_or_created_just_overwrites(self, path):
        Table.read(path, "T").add_row("other", [0, 0, 0])
        stranger = Table(
            data=[["only"] * 3], row_labels=["mine"], column_headers=["a", "b", "c"], name="T"
        )
        # `stranger` was built from scratch, so it has no base: it replaces the table outright
        stranger.write(path)
        assert stranger.last_merge is None
        assert written(path).data.row_labels == ["mine"]

    def test_a_created_table_has_a_base_from_then_on(self, path):
        t = Table(data=[[1]], row_labels=["r"], column_headers=["v"], name="M")
        t.create(path, sheet="Data")
        other = Table.read(path, "M")
        other.add_row("theirs", [2])
        quiet_write(other, path)
        t.add_row("ours", [3])
        quiet_write(t, path)
        assert written_named(path, "M").data.row_labels == ["r", "theirs", "ours"]


def written_named(path, name):
    return Table.read(path, name)


class TestDuplicateLabels:
    def _duplicate_table(self, path):
        """Force a table with a repeated row label onto disk, as a hand-edit could."""
        wb = load_workbook(path)
        ws = wb["Data"]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "r3":
                    cell.value = "r1"
        wb.save(path)

    def test_ours_wins_wholesale_with_a_warning(self, path):
        alice = Table.read(path, "T")
        self._duplicate_table(path)
        alice.set_cell(row="r1", column="a", value="alice")
        with pytest.warns(MergeConflictWarning, match="not unique"):
            alice.write(path)
        assert written(path).read_cell(row="r1", column="a") == "alice"


class TestManyConflictsAreTruncated:
    def test_the_warning_lists_ten_and_counts_the_rest(self, book):
        create_sheet(book, "Data")
        n = 25
        Table(
            data=[[0]] * n,
            row_labels=[f"r{i}" for i in range(n)],
            column_headers=["v"],
            name="Big",
        ).create(book, sheet="Data")
        alice, bob = Table.read(book, "Big"), Table.read(book, "Big")
        for i in range(n):
            alice.set_cell(row=f"r{i}", column="v", value="alice")
            bob.set_cell(row=f"r{i}", column="v", value="bob")
        quiet_write(alice, book)
        with pytest.warns(MergeConflictWarning) as caught:
            bob.write(book)
        message = str(caught[0].message)
        assert "(25)" in message
        assert "and 15 more" in message
        assert message.count("cell (") == 10
        assert len(bob.last_merge.conflicts) == 25  # the report keeps them all


class TestDeletedUnderneath:
    def test_writing_a_deleted_table_raises_and_leaves_the_object_alone(self, path, users):
        _, bob = users
        delete_table(path, "T")
        bob.add_row("bob", [0, 0, 0])
        with pytest.raises(TableNotFoundError):
            bob.write(path)
        assert bob.data.row_labels == ["r1", "r2", "r3", "bob"]


class TestWarningLocation:
    def test_the_warning_points_at_the_callers_write_call(self, path, users):
        alice, bob = users
        alice.set_cell(row="r1", column="a", value="alice")
        bob.set_cell(row="r1", column="a", value="bob")
        quiet_write(alice, path)
        with pytest.warns(MergeConflictWarning) as caught:
            bob.write(path)
        assert caught[0].filename == __file__
