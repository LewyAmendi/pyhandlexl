"""Unit tests for the three-way merge engine, with no files involved."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta

import pytest

from pyhandlexl import ColumnType, TableStyle
from pyhandlexl._merge import NOT_UNIQUE, MergeReport, Snapshot, merge, same_state, same_value

ANY = ColumnType.ANY


def snap(rows, cols, data, corner="", style=TableStyle.DEFAULT, types=None):
    return Snapshot(list(rows), list(cols), [list(r) for r in data], corner, style,
                    list(types) if types is not None else [ANY] * len(cols))  # fmt: skip


def identity(labels):
    return {label: label for label in labels}


def run(base, ours, theirs, row_origin=None, col_origin=None):
    """Merge, with origins defaulting to 'nothing renamed, nothing added'."""
    if row_origin is None:
        row_origin = {r: (r if r in base.row_labels else None) for r in ours.row_labels}
    if col_origin is None:
        col_origin = {c: (c if c in base.column_headers else None) for c in ours.column_headers}
    return merge(base, ours, row_origin, col_origin, theirs)


@pytest.fixture
def base():
    return snap(["r1", "r2", "r3"], ["a", "b"], [[1, 2], [3, 4], [5, 6]])


class TestSameValue:
    @pytest.mark.parametrize(
        ("a", "b"),
        [
            (1, 1),
            (5, 5.0),
            ("x", "x"),
            (None, None),
            (True, True),
            (date(2026, 1, 2), datetime(2026, 1, 2)),
            (datetime(2026, 1, 2, 3, 4, 5, 123456), datetime(2026, 1, 2, 3, 4, 5, 123000)),
            (time(1, 2, 3, 500), time(1, 2, 3, 0)),
            (timedelta(seconds=5), timedelta(seconds=5, microseconds=400)),
        ],
    )
    def test_equal(self, a, b):
        assert same_value(a, b)

    @pytest.mark.parametrize(
        ("a", "b"),
        [
            (True, 1),
            (False, 0),
            (1, "1"),
            (None, ""),
            (None, 0),
            ("a", "A"),
            (1, 2),
            (datetime(2026, 1, 2, 3), datetime(2026, 1, 2, 4)),
            (datetime(2026, 1, 2, 0, 0, 0, 5000), datetime(2026, 1, 2)),
            (time(1, 0), time(2, 0)),
            (timedelta(seconds=1), timedelta(seconds=2)),
            (date(2026, 1, 2), time(0, 0)),
        ],
    )
    def test_different(self, a, b):
        assert not same_value(a, b)

    def test_nan_is_not_equal_to_itself_but_identity_short_circuits(self):
        nan = float("nan")
        assert same_value(nan, nan)  # same object


class TestSameState:
    def test_equal_snapshots(self, base):
        assert same_state(base, base.copy())

    @pytest.mark.parametrize(
        "change",
        [
            lambda s: s.data[0].__setitem__(0, 99),
            lambda s: s.row_labels.__setitem__(0, "z"),
            lambda s: s.column_headers.__setitem__(0, "z"),
            lambda s: setattr(s, "corner", "c"),
            lambda s: setattr(s, "style", TableStyle.MINIMAL),
            lambda s: s.column_types.__setitem__(0, ColumnType.NUMBER),
        ],
    )
    def test_any_difference_is_seen(self, base, change):
        other = base.copy()
        change(other)
        assert not same_state(base, other)

    def test_excel_round_trip_noise_is_not_a_difference(self, base):
        other = base.copy()
        other.data[0][0] = 1.0
        assert same_state(base, other)

    def test_copy_is_independent(self, base):
        other = base.copy()
        other.data[0][0] = 99
        other.row_labels.append("x")
        assert base.data[0][0] == 1
        assert base.row_labels == ["r1", "r2", "r3"]


class TestNothingToReconcile:
    def test_theirs_equals_base_keeps_ours(self, base):
        ours = base.copy()
        ours.data[0][0] = 10
        merged, report = run(base, ours, base.copy())
        assert merged.data == ours.data
        assert report == MergeReport()

    def test_ours_equals_base_takes_theirs(self, base):
        theirs = base.copy()
        theirs.data[1][1] = 40
        merged, report = run(base, base.copy(), theirs)
        assert merged.data[1][1] == 40
        assert report.cells_updated == 1
        assert not report.conflicts


class TestCells:
    def test_edits_to_different_cells_both_survive(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = "mine"
        theirs.data[2][1] = "theirs"
        merged, report = run(base, ours, theirs)
        assert merged.data == [["mine", 2], [3, 4], [5, "theirs"]]
        assert not report.conflicts
        assert report.cells_updated == 1

    def test_same_cell_same_value_is_not_a_conflict(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = theirs.data[0][0] = 7
        merged, report = run(base, ours, theirs)
        assert merged.data[0][0] == 7
        assert not report.conflicts

    def test_same_cell_different_values_ours_wins_and_is_reported(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = "mine"
        theirs.data[0][0] = "theirs"
        merged, report = run(base, ours, theirs)
        assert merged.data[0][0] == "mine"
        assert report.conflicts == (
            "cell ('r1', 'a'): yours 'mine' overwrote another writer's 'theirs'",
        )

    def test_our_edit_back_to_the_base_value_yields_to_theirs(self, base):
        # "we left it alone" is judged by value, not by intent
        ours, theirs = base.copy(), base.copy()
        theirs.data[0][0] = 100
        merged, _ = run(base, ours, theirs)
        assert merged.data[0][0] == 100

    def test_clearing_a_cell_is_an_edit(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = None
        theirs.data[0][0] = 100
        merged, report = run(base, ours, theirs)
        assert merged.data[0][0] is None
        assert len(report.conflicts) == 1

    def test_long_values_are_shortened_in_the_message(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = "x" * 100
        theirs.data[0][0] = "y" * 100
        _, report = run(base, ours, theirs)
        assert "..." in report.conflicts[0]
        assert len(report.conflicts[0]) < 140

    def test_theirs_touching_a_cell_we_never_changed_is_not_a_conflict_even_if_bool_vs_int(
        self, base
    ):
        theirs = base.copy()
        theirs.data[0][0] = True  # 1 -> True is a real change
        merged, report = run(base, base.copy(), theirs)
        assert merged.data[0][0] is True
        assert report.cells_updated == 1


class TestRows:
    def test_rows_added_by_both_sides_all_appear(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.append("mine")
        ours.data.append([7, 7])
        theirs.row_labels.append("theirs")
        theirs.data.append([8, 8])
        merged, report = run(base, ours, theirs)
        assert set(merged.row_labels) == {"r1", "r2", "r3", "mine", "theirs"}
        assert merged.data[merged.row_labels.index("mine")] == [7, 7]
        assert merged.data[merged.row_labels.index("theirs")] == [8, 8]
        assert report.rows_added == ("theirs",)
        assert not report.conflicts

    def test_both_add_the_same_label_with_the_same_data_is_no_conflict(self, base):
        ours, theirs = base.copy(), base.copy()
        for s in (ours, theirs):
            s.row_labels.append("new")
            s.data.append([9, 9])
        merged, report = run(base, ours, theirs)
        assert merged.row_labels.count("new") == 1
        assert not report.conflicts

    def test_both_add_the_same_label_with_different_data_ours_wins_cell_by_cell(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.append("new")
        ours.data.append([1, 1])
        theirs.row_labels.append("new")
        theirs.data.append([2, 2])
        merged, report = run(base, ours, theirs)
        assert merged.data[merged.row_labels.index("new")] == [1, 1]
        assert len(report.conflicts) == 2

    def test_their_delete_removes_an_untouched_row(self, base):
        theirs = base.copy()
        del theirs.row_labels[1], theirs.data[1]
        merged, report = run(base, base.copy(), theirs)
        assert merged.row_labels == ["r1", "r3"]
        assert report.rows_removed == ("r2",)
        assert not report.conflicts

    def test_their_delete_of_a_row_we_edited_is_undone_and_reported(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[1][0] = "edited"
        del theirs.row_labels[1], theirs.data[1]
        merged, report = run(base, ours, theirs)
        assert "r2" in merged.row_labels
        assert merged.data[merged.row_labels.index("r2")][0] == "edited"
        assert any("deleted by another writer but edited by you" in c for c in report.conflicts)

    def test_our_delete_of_a_row_they_edited_wins_and_is_reported(self, base):
        ours, theirs = base.copy(), base.copy()
        del ours.row_labels[1], ours.data[1]
        theirs.data[1][0] = "theirs edit"
        merged, report = run(base, ours, theirs)
        assert "r2" not in merged.row_labels
        assert any("deleted by you but edited by another writer" in c for c in report.conflicts)

    def test_our_delete_of_a_row_they_left_alone_is_silent(self, base):
        ours, theirs = base.copy(), base.copy()
        del ours.row_labels[1], ours.data[1]
        theirs.data[0][0] = 100  # some other row
        merged, report = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "r3"]
        assert merged.data[0][0] == 100
        assert not report.conflicts

    def test_both_delete_the_same_row(self, base):
        ours, theirs = base.copy(), base.copy()
        for s in (ours, theirs):
            del s.row_labels[1], s.data[1]
        merged, report = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "r3"]
        assert not report.conflicts

    def test_their_added_cells_in_a_column_we_dont_have_count_as_an_edit_to_a_deleted_row(
        self, base
    ):
        # they added a column c and filled it for r2; we deleted r2 -> their edit is discarded
        ours, theirs = base.copy(), base.copy()
        del ours.row_labels[1], ours.data[1]
        theirs.column_headers.append("c")
        theirs.column_types.append(ANY)
        for row in theirs.data:
            row.append(None)
        theirs.data[1][2] = "filled"
        _, report = run(base, ours, theirs)
        assert any("deleted by you but edited" in c for c in report.conflicts)

    def test_their_appended_row_lands_where_they_put_it(self, base):
        theirs = base.copy()
        theirs.row_labels.append("t")
        theirs.data.append([0, 0])
        merged, _ = run(base, base.copy(), theirs)
        assert merged.row_labels == ["r1", "r2", "r3", "t"]


class TestRowOrder:
    def test_our_append_goes_after_their_append(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.append("mine")
        ours.data.append([7, 7])
        theirs.row_labels.append("theirs")
        theirs.data.append([8, 8])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "r2", "r3", "theirs", "mine"]

    def test_several_appends_of_ours_keep_their_order_after_theirs(self, base):
        ours, theirs = base.copy(), base.copy()
        for label in ("m1", "m2", "m3"):
            ours.row_labels.append(label)
            ours.data.append([0, 0])
        theirs.row_labels.append("theirs")
        theirs.data.append([8, 8])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "r2", "r3", "theirs", "m1", "m2", "m3"]

    def test_appending_to_a_table_they_also_appended_to_when_it_started_empty(self):
        base = snap([], ["a"], [])
        ours = snap(["mine"], ["a"], [[1]])
        theirs = snap(["theirs"], ["a"], [[2]])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["theirs", "mine"]

    def test_a_mid_table_insert_still_follows_its_predecessor_when_they_appended(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.insert(1, "mine")
        ours.data.insert(1, [0, 0])
        theirs.row_labels.append("theirs")
        theirs.data.append([8, 8])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "mine", "r2", "r3", "theirs"]

    def test_appended_columns_follow_the_same_rule(self, base):
        ours, theirs = base.copy(), base.copy()
        for s, header in ((ours, "mine"), (theirs, "theirs")):
            s.column_headers.append(header)
            s.column_types.append(ANY)
            for row in s.data:
                row.append(header)
        merged, _ = run(base, ours, theirs)
        assert merged.column_headers == ["a", "b", "theirs", "mine"]

    def test_our_added_row_follows_its_predecessor(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.insert(1, "mine")
        ours.data.insert(1, [0, 0])
        theirs.row_labels.append("theirs")
        theirs.data.append([9, 9])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "mine", "r2", "r3", "theirs"]

    def test_our_row_inserted_first_stays_first(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.insert(0, "first")
        ours.data.insert(0, [0, 0])
        theirs.row_labels.insert(0, "theirs-first")
        theirs.data.insert(0, [1, 1])
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels[0] == "first"
        assert "theirs-first" in merged.row_labels

    def test_predecessor_deleted_by_them_falls_back_to_the_previous_survivor(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.insert(2, "mine")  # r1 r2 mine r3
        ours.data.insert(2, [0, 0])
        del theirs.row_labels[1], theirs.data[1]  # they delete r2
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels == ["r1", "mine", "r3"]

    def test_all_predecessors_gone_goes_to_the_end(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.append("mine")  # after r3
        ours.data.append([0, 0])
        theirs.row_labels[:] = ["only"]
        theirs.data[:] = [[0, 0]]
        merged, _ = run(base, ours, theirs)
        assert merged.row_labels[-1] == "mine"
        assert set(merged.row_labels) == {"only", "mine"}

    def test_their_reorder_is_adopted_for_rows_we_did_not_touch(self, base):
        theirs = base.copy()
        theirs.row_labels[:] = ["r3", "r1", "r2"]
        theirs.data[:] = [[5, 6], [1, 2], [3, 4]]
        merged, _ = run(base, base.copy(), theirs)
        assert merged.row_labels == ["r3", "r1", "r2"]


class TestRenames:
    def test_our_rename_keeps_their_edit_of_the_same_row(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels[0] = "renamed"
        theirs.data[0][0] = "theirs"
        merged, report = run(
            base, ours, theirs, row_origin={"renamed": "r1", "r2": "r2", "r3": "r3"}
        )
        assert merged.row_labels[0] == "renamed"
        assert merged.data[0][0] == "theirs"
        assert not report.conflicts

    def test_our_rename_is_not_mistaken_for_delete_plus_add(self, base):
        ours = base.copy()
        ours.row_labels[0] = "renamed"
        theirs = base.copy()
        theirs.data[2][0] = "elsewhere"
        merged, report = run(
            base, ours, theirs, row_origin={"renamed": "r1", "r2": "r2", "r3": "r3"}
        )
        assert merged.row_labels == ["renamed", "r2", "r3"]
        assert report.rows_added == ()
        assert report.rows_removed == ()

    def test_their_rename_arrives_as_delete_plus_add_when_we_left_the_row_alone(self, base):
        theirs = base.copy()
        theirs.row_labels[0] = "theirs-name"
        merged, report = run(base, base.copy(), theirs)
        assert merged.row_labels[0] == "theirs-name"
        assert "r1" not in merged.row_labels
        assert report.rows_added == ("theirs-name",)
        assert report.rows_removed == ("r1",)

    def test_renaming_onto_a_label_they_just_added_ours_wins(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels[0] = "clash"
        theirs.row_labels.append("clash")
        theirs.data.append([0, 0])
        merged, report = run(base, ours, theirs, row_origin={"clash": "r1", "r2": "r2", "r3": "r3"})
        assert merged.row_labels.count("clash") == 1
        assert merged.data[merged.row_labels.index("clash")] == [1, 2]  # our row's data
        assert any("you renamed one to" in c for c in report.conflicts)

    def test_column_rename_carries_their_edits_in_that_column(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.column_headers[0] = "alpha"
        theirs.data[0][0] = "theirs"
        merged, report = run(base, ours, theirs, col_origin={"alpha": "a", "b": "b"})
        assert merged.column_headers == ["alpha", "b"]
        assert merged.data[0][0] == "theirs"
        assert not report.conflicts


class TestColumns:
    def test_columns_added_by_both_sides_all_appear(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.column_headers.append("mine")
        ours.column_types.append(ANY)
        for row in ours.data:
            row.append("m")
        theirs.column_headers.append("theirs")
        theirs.column_types.append(ANY)
        for row in theirs.data:
            row.append("t")
        merged, report = run(base, ours, theirs)
        assert set(merged.column_headers) == {"a", "b", "mine", "theirs"}
        j, k = merged.column_headers.index("mine"), merged.column_headers.index("theirs")
        assert all(row[j] == "m" and row[k] == "t" for row in merged.data)
        assert report.columns_added == ("theirs",)

    def test_a_column_added_to_rows_we_also_added_is_blank_where_neither_side_has_it(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.row_labels.append("mine")
        ours.data.append([1, 1])
        theirs.column_headers.append("extra")
        theirs.column_types.append(ANY)
        for row in theirs.data:
            row.append("x")
        merged, _ = run(base, ours, theirs)
        j = merged.column_headers.index("extra")
        assert merged.data[merged.row_labels.index("mine")][j] is None
        assert merged.data[0][j] == "x"

    def test_their_column_drop_removes_an_untouched_column(self, base):
        theirs = base.copy()
        theirs.column_headers.pop(1)
        theirs.column_types.pop(1)
        for row in theirs.data:
            row.pop(1)
        merged, report = run(base, base.copy(), theirs)
        assert merged.column_headers == ["a"]
        assert report.columns_removed == ("b",)
        assert merged.data == [[1], [3], [5]]

    def test_their_drop_of_a_column_we_edited_is_undone_and_reported(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][1] = "edited"
        theirs.column_headers.pop(1)
        theirs.column_types.pop(1)
        for row in theirs.data:
            row.pop(1)
        merged, report = run(base, ours, theirs)
        assert merged.column_headers == ["a", "b"]
        assert merged.data[0][1] == "edited"
        assert any("deleted by another writer but edited by you" in c for c in report.conflicts)

    def test_both_dropping_different_columns_that_leaves_none_falls_back_to_ours(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.column_headers[:] = ["b"]
        ours.column_types[:] = [ANY]
        ours.data[:] = [[2], [4], [6]]
        theirs.column_headers[:] = ["a"]
        theirs.column_types[:] = [ANY]
        theirs.data[:] = [[1], [3], [5]]
        merged, report = run(base, ours, theirs)
        assert merged.column_headers == ["b"]
        assert merged.data == ours.data
        assert "no columns" in report.conflicts[0]


class TestCornerStyleAndTypes:
    def test_corner_changed_only_by_them_is_taken(self, base):
        theirs = base.copy()
        theirs.corner = "their corner"
        merged, report = run(base, base.copy(), theirs)
        assert merged.corner == "their corner"
        assert not report.conflicts

    def test_corner_changed_only_by_us_is_kept(self, base):
        ours = base.copy()
        ours.corner = "mine"
        merged, _ = run(base, ours, base.copy())
        assert merged.corner == "mine"

    def test_corner_changed_by_both_differently_ours_wins(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.corner, theirs.corner = "mine", "theirs"
        merged, report = run(base, ours, theirs)
        assert merged.corner == "mine"
        assert report.conflicts == ("corner: yours 'mine' overwrote 'theirs'",)

    def test_corner_changed_to_the_same_thing_is_no_conflict(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.corner = theirs.corner = "same"
        _, report = run(base, ours, theirs)
        assert not report.conflicts

    def test_style_changed_only_by_them_is_taken(self, base):
        theirs = base.copy()
        theirs.style = TableStyle.MINIMAL
        merged, report = run(base, base.copy(), theirs)
        assert merged.style == TableStyle.MINIMAL
        assert not report.conflicts

    def test_style_changed_only_by_us_is_kept(self, base):
        ours = base.copy()
        ours.style = TableStyle.NONE
        merged, _ = run(base, ours, base.copy())
        assert merged.style == TableStyle.NONE

    def test_style_changed_by_both_differently_ours_wins(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.style, theirs.style = TableStyle.NONE, TableStyle.MINIMAL
        merged, report = run(base, ours, theirs)
        assert merged.style == TableStyle.NONE
        assert any(c.startswith("style:") for c in report.conflicts)

    def test_column_type_changed_only_by_them_is_taken(self, base):
        theirs = base.copy()
        theirs.column_types[0] = ColumnType.NUMBER
        merged, report = run(base, base.copy(), theirs)
        assert merged.column_types == [ColumnType.NUMBER, ANY]
        assert not report.conflicts

    def test_column_type_changed_only_by_us_is_kept(self, base):
        ours = base.copy()
        ours.column_types[1] = ColumnType.TEXT
        merged, _ = run(base, ours, base.copy())
        assert merged.column_types == [ANY, ColumnType.TEXT]

    def test_column_type_changed_by_both_differently_ours_wins(self, base):
        ours, theirs = base.copy(), base.copy()
        ours.column_types[0] = ColumnType.TEXT
        theirs.column_types[0] = ColumnType.NUMBER
        merged, report = run(base, ours, theirs)
        assert merged.column_types[0] == ColumnType.TEXT
        assert len(report.conflicts) == 1

    def test_a_column_both_added_with_different_types_ours_wins(self, base):
        ours, theirs = base.copy(), base.copy()
        for s, kind in ((ours, ColumnType.TEXT), (theirs, ColumnType.NUMBER)):
            s.column_headers.append("new")
            s.column_types.append(kind)
            for row in s.data:
                row.append(None)
        merged, report = run(base, ours, theirs)
        assert merged.column_types[merged.column_headers.index("new")] == ColumnType.TEXT
        assert len(report.conflicts) == 1

    def test_types_travel_with_their_columns_through_a_reorder(self, base):
        theirs = base.copy()
        theirs.column_types[1] = ColumnType.NUMBER
        theirs.column_headers[:] = ["b", "a"]
        theirs.column_types[:] = [ColumnType.NUMBER, ANY]
        theirs.data[:] = [[2, 1], [4, 3], [6, 5]]
        merged, _ = run(base, base.copy(), theirs)
        assert merged.column_headers == ["b", "a"]
        assert merged.column_types == [ColumnType.NUMBER, ANY]


class TestFallbacks:
    @pytest.mark.parametrize("which", ["base", "ours", "theirs"])
    def test_duplicate_labels_anywhere_means_ours_wins_wholesale(self, base, which):
        ours, theirs = base.copy(), base.copy()
        ours.data[0][0] = "mine"
        theirs.data[1][1] = "theirs"
        target = {"base": base, "ours": ours, "theirs": theirs}[which]
        target.row_labels[2] = "r1"
        merged, report = run(base, ours, theirs)
        assert merged.data == ours.data
        assert merged.row_labels == ours.row_labels
        assert report.conflicts == (NOT_UNIQUE,)

    def test_duplicate_headers_also_fall_back(self, base):
        theirs = base.copy()
        theirs.column_headers[1] = "a"
        _, report = run(base, base.copy(), theirs)
        assert report.conflicts == (NOT_UNIQUE,)

    def test_the_result_is_a_copy_not_the_input(self, base):
        theirs = base.copy()
        theirs.row_labels[2] = "r1"
        ours = base.copy()
        merged, _ = run(base, ours, theirs)
        merged.data[0][0] = "mutated"
        assert ours.data[0][0] == 1


class TestZeroRows:
    def test_merging_into_an_empty_table(self):
        base = snap([], ["a"], [])
        theirs = snap(["r"], ["a"], [[1]])
        merged, report = run(base, base.copy(), theirs)
        assert merged.row_labels == ["r"]
        assert report.rows_added == ("r",)

    def test_we_add_rows_to_a_table_they_emptied(self):
        base = snap(["r"], ["a"], [[1]])
        ours = snap(["r", "n"], ["a"], [[1], [2]])
        theirs = snap([], ["a"], [])
        merged, report = run(base, ours, theirs)
        assert merged.row_labels == ["n"]
        assert report.rows_removed == ("r",)


class TestReport:
    def test_defaults_are_empty(self):
        report = MergeReport()
        assert report.rows_added == report.rows_removed == ()
        assert report.columns_added == report.columns_removed == ()
        assert report.cells_updated == 0
        assert report.conflicts == ()

    def test_is_frozen(self):
        with pytest.raises(AttributeError):
            MergeReport().cells_updated = 3  # type: ignore[misc]

    def test_it_is_a_public_name_of_the_package_not_of_a_private_module(self):
        import pickle

        import pyhandlexl

        assert pyhandlexl.MergeReport is MergeReport
        assert MergeReport.__module__ == "pyhandlexl"
        assert repr(MergeReport()).startswith("MergeReport(")
        report = MergeReport(rows_added=("a",), conflicts=("cell b/c",))
        assert pickle.loads(pickle.dumps(report)) == report
