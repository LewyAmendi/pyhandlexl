"""Tests for the grid editing helpers."""

from __future__ import annotations

import pytest

from pyhandlexl import grid


@pytest.fixture
def g():
    return [["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]]


class TestReturnsNewGrid:
    def test_input_is_not_modified(self, g):
        original = [row[:] for row in g]
        grid.set_value(g, 1, 1, "X")
        grid.delete_row(g, 2)
        grid.append_column(g, [1, 2, 3])
        assert g == original


class TestReading:
    def test_get_row(self, g):
        assert grid.get_row(g, 1) == ["a", "b", "c"]
        assert grid.get_row(g, 3) == ["g", "h", "i"]

    def test_get_column(self, g):
        assert grid.get_column(g, 2) == ["b", "e", "h"]

    def test_get_column_pads_short_rows(self):
        assert grid.get_column([["a"], ["b", "c"]], 2) == [None, "c"]

    def test_out_of_range_raises(self, g):
        with pytest.raises(IndexError):
            grid.get_row(g, 0)
        with pytest.raises(IndexError):
            grid.get_row(g, 4)
        with pytest.raises(IndexError):
            grid.get_column(g, 4)


class TestSetValue:
    def test_sets_one_cell(self, g):
        assert grid.set_value(g, 2, 3, "Z")[1] == ["d", "e", "Z"]

    def test_pads_a_short_row_to_reach_the_column(self):
        out = grid.set_value([["a", "b", "c"], ["d"]], 2, 3, "Z")
        assert out == [["a", "b", "c"], ["d", None, "Z"]]


class TestSetLine:
    def test_set_row(self, g):
        assert grid.set_row(g, 1, [1, 2])[0] == [1, 2]

    def test_set_column(self, g):
        assert grid.set_column(g, 1, ["x", "y", "z"]) == [
            ["x", "b", "c"],
            ["y", "e", "f"],
            ["z", "h", "i"],
        ]

    def test_set_column_wrong_length_raises(self, g):
        with pytest.raises(ValueError):
            grid.set_column(g, 1, ["x", "y"])


class TestInsert:
    def test_insert_row_in_the_middle(self, g):
        out = grid.insert_row(g, 2, ["N", "N", "N"])
        assert out[1] == ["N", "N", "N"]
        assert len(out) == 4

    def test_insert_row_at_the_end(self, g):
        assert grid.insert_row(g, 4, ["N"])[3] == ["N"]

    def test_insert_row_past_the_end_raises(self, g):
        with pytest.raises(IndexError):
            grid.insert_row(g, 5, ["N"])

    def test_insert_column(self, g):
        out = grid.insert_column(g, 2, ["1", "2", "3"])
        assert out[0] == ["a", "1", "b", "c"]

    def test_insert_column_wrong_length_raises(self, g):
        with pytest.raises(ValueError):
            grid.insert_column(g, 1, ["1"])


class TestAppend:
    def test_append_row(self, g):
        out = grid.append_row(g, ["j", "k", "l"])
        assert out[-1] == ["j", "k", "l"]
        assert len(out) == 4

    def test_append_column(self, g):
        out = grid.append_column(g, [1, 2, 3])
        assert [row[-1] for row in out] == [1, 2, 3]

    def test_append_column_pads_ragged_rows_first(self):
        out = grid.append_column([["a", "b"], ["c"]], ["x", "y"])
        assert out == [["a", "b", "x"], ["c", None, "y"]]

    def test_append_column_wrong_length_raises(self, g):
        with pytest.raises(ValueError):
            grid.append_column(g, [1, 2])


class TestDelete:
    def test_delete_row(self, g):
        assert grid.delete_row(g, 2) == [["a", "b", "c"], ["g", "h", "i"]]

    def test_delete_column(self, g):
        assert grid.delete_column(g, 2) == [["a", "c"], ["d", "f"], ["g", "i"]]

    def test_delete_out_of_range_raises(self, g):
        with pytest.raises(IndexError):
            grid.delete_row(g, 9)
        with pytest.raises(IndexError):
            grid.delete_column(g, 9)


class TestShape:
    def test_pad(self):
        assert grid.pad([["a"], ["b", "c", "d"]]) == [["a", None, None], ["b", "c", "d"]]

    def test_transpose(self, g):
        assert grid.transpose(g) == [
            ["a", "d", "g"],
            ["b", "e", "h"],
            ["c", "f", "i"],
        ]

    def test_transpose_pads_ragged_rows(self):
        assert grid.transpose([["a", "b"], ["c"]]) == [["a", "c"], ["b", None]]

    def test_transpose_empty(self):
        assert grid.transpose([]) == []
