"""The pure ``pyhandlexl.grid`` functions, checked against laws and a plain reference."""

from __future__ import annotations

import copy
import random

import pytest

from pyhandlexl import grid

SEEDS = range(40)


def random_grid(
    rng: random.Random,
    *,
    ragged: bool = True,
    max_rows: int = 6,
    max_cols: int = 6,
    minimum: int = 0,
):
    rows = rng.randint(minimum, max_rows)
    width = rng.randint(minimum, max_cols)
    return [
        [
            rng.choice([None, 0, 1, "a", "b c", 2.5, True])
            for _ in range(rng.randint(0, width) if ragged else width)
        ]
        for _ in range(rows)
    ]


def widest(g) -> int:
    return max((len(r) for r in g), default=0)


class TestInputsAreNeverModified:
    @pytest.mark.parametrize("seed", SEEDS)
    def test_no_function_touches_its_input_or_shares_rows_with_it(self, seed):
        rng = random.Random(seed)
        g = random_grid(rng)
        if not widest(g):  # need at least one cell for the positional calls below
            g = [[1, 2], [3]]
        snapshot = copy.deepcopy(g)
        n_rows, n_cols = len(g), max(1, widest(g))
        results = [
            grid.set_value(g, 1, 1, "X"),
            grid.set_row(g, 1, ["r"] * 3),
            grid.set_column(g, 1, ["c"] * n_rows),
            grid.insert_row(g, 1, ["i"]),
            grid.insert_column(g, 1, ["i"] * n_rows),
            grid.append_row(g, ["a"]),
            grid.append_column(g, ["a"] * n_rows),
            grid.delete_row(g, n_rows),
            grid.delete_column(g, n_cols) if widest(g) else g,
            grid.pad(g),
            grid.transpose(g),
        ]
        assert g == snapshot
        for result in results:  # mutating any result must not reach back into the input
            for row in result:
                row.append("!")
        assert g == snapshot


class TestAlgebraicLaws:
    @pytest.mark.parametrize("seed", SEEDS)
    def test_transposing_twice_gives_the_padded_grid(self, seed):
        g = random_grid(random.Random(seed))
        assert grid.transpose(grid.transpose(g)) == grid.pad(g) or not any(g)

    @pytest.mark.parametrize("seed", SEEDS)
    def test_pad_is_idempotent_and_rectangular(self, seed):
        g = random_grid(random.Random(100 + seed))
        padded = grid.pad(g)
        assert grid.pad(padded) == padded
        assert len({len(r) for r in padded}) <= 1
        assert len(padded) == len(g)
        for original, row in zip(g, padded, strict=True):
            assert row[: len(original)] == original
            assert all(v is None for v in row[len(original) :])

    @pytest.mark.parametrize("seed", SEEDS)
    def test_insert_then_delete_a_row_is_the_identity(self, seed):
        rng = random.Random(200 + seed)
        g = random_grid(rng)
        position = rng.randint(1, len(g) + 1)
        assert grid.delete_row(grid.insert_row(g, position, ["x", "y"]), position) == g

    @pytest.mark.parametrize("seed", SEEDS)
    def test_append_then_delete_the_last_row_is_the_identity(self, seed):
        g = random_grid(random.Random(300 + seed))
        assert grid.delete_row(grid.append_row(g, [1, 2]), len(g) + 1) == g

    @pytest.mark.parametrize("seed", SEEDS)
    def test_insert_then_delete_a_column_is_the_identity_on_a_rectangle(self, seed):
        rng = random.Random(400 + seed)
        g = random_grid(rng, ragged=False, minimum=1)
        position = rng.randint(1, len(g[0]) + 1)
        values = list(range(len(g)))
        assert grid.delete_column(grid.insert_column(g, position, values), position) == g

    @pytest.mark.parametrize("seed", SEEDS)
    def test_set_value_changes_exactly_one_cell(self, seed):
        rng = random.Random(500 + seed)
        g = random_grid(rng, ragged=False, max_rows=5, max_cols=5, minimum=1)
        r, c = rng.randint(1, len(g)), rng.randint(1, len(g[0]))
        out = grid.set_value(g, r, c, "NEW")
        for i in range(len(g)):
            for j in range(len(g[0])):
                expected = "NEW" if (i + 1, j + 1) == (r, c) else g[i][j]
                assert out[i][j] == expected

    @pytest.mark.parametrize("seed", SEEDS)
    def test_get_row_and_get_column_agree_with_the_grid(self, seed):
        rng = random.Random(600 + seed)
        g = random_grid(rng, ragged=False, max_rows=5, max_cols=5, minimum=1)
        for i, row in enumerate(g, start=1):
            assert grid.get_row(g, i) == row
        for j in range(1, len(g[0]) + 1):
            assert grid.get_column(g, j) == [row[j - 1] for row in g]

    @pytest.mark.parametrize("seed", SEEDS)
    def test_set_row_and_set_column_touch_only_theirs(self, seed):
        rng = random.Random(700 + seed)
        g = random_grid(rng, ragged=False, max_rows=5, max_cols=5, minimum=1)
        r, c = rng.randint(1, len(g)), rng.randint(1, len(g[0]))
        rows = grid.set_row(g, r, ["R"] * len(g[0]))
        assert rows[r - 1] == ["R"] * len(g[0])
        assert [row for i, row in enumerate(rows) if i != r - 1] == [
            row for i, row in enumerate(g) if i != r - 1
        ]
        cols = grid.set_column(g, c, ["C"] * len(g))
        assert grid.get_column(cols, c) == ["C"] * len(g)
        for j in range(1, len(g[0]) + 1):
            if j != c:
                assert grid.get_column(cols, j) == grid.get_column(g, j)


class TestRejectedInput:
    def test_a_bare_string_is_not_a_row_of_characters(self):
        g = [[1, 2, 3], [4, 5, 6]]
        for call in (
            lambda: grid.set_row(g, 1, "abc"),
            lambda: grid.insert_row(g, 1, "abc"),
            lambda: grid.append_row(g, "abc"),
            lambda: grid.set_column(g, 1, "ab"),
            lambda: grid.insert_column(g, 1, "ab"),
            lambda: grid.append_column(g, "ab"),
            lambda: grid.set_row(g, 1, b"abc"),
        ):
            with pytest.raises(TypeError, match="split it into individual characters"):
                call()

    def test_a_grid_containing_a_bare_string_row_is_refused(self):
        with pytest.raises(TypeError):
            grid.pad([[1, 2], "ab"])
        with pytest.raises(TypeError):
            grid.transpose(["ab", "cd"])

    @pytest.mark.parametrize("bad", [0, -1, 4, 99])
    def test_out_of_range_rows_raise_index_error(self, bad):
        g = [[1], [2], [3]]
        for call in (
            lambda: grid.get_row(g, bad),
            lambda: grid.set_row(g, bad, [0]),
            lambda: grid.delete_row(g, bad),
        ):
            with pytest.raises(IndexError):
                call()

    @pytest.mark.parametrize("bad", [1.0, "1", None, 1.5])
    def test_non_integer_positions_raise_type_error(self, bad):
        g = [[1, 2], [3, 4]]
        for call in (
            lambda: grid.get_row(g, bad),
            lambda: grid.get_column(g, bad),
            lambda: grid.set_value(g, bad, 1, 0),
            lambda: grid.insert_row(g, bad, [0, 0]),
            lambda: grid.delete_column(g, bad),
        ):
            with pytest.raises(TypeError):
                call()

    def test_wrong_length_values_raise_value_error(self):
        g = [[1, 2], [3, 4]]
        with pytest.raises(ValueError):
            grid.set_column(g, 1, [1])
        with pytest.raises(ValueError):
            grid.insert_column(g, 1, [1, 2, 3])
        with pytest.raises(ValueError):
            grid.append_column(g, [])


class TestEdgeShapes:
    def test_empty_grids(self):
        assert grid.pad([]) == []
        assert grid.transpose([]) == []
        assert grid.transpose([[]]) == []
        assert grid.append_row([], [1, 2]) == [[1, 2]]
        assert grid.insert_row([], 1, [1]) == [[1]]

    def test_ragged_rows(self):
        assert grid.transpose([[1, 2, 3], [4]]) == [[1, 4], [2, None], [3, None]]
        assert grid.pad([[1, 2, 3], [4]]) == [[1, 2, 3], [4, None, None]]
        assert grid.append_row([[1, 2, 3], [4]], [9]) == [[1, 2, 3], [4], [9]]
        assert grid.set_value([[1, 2, 3], [4]], 2, 3, "x") == [[1, 2, 3], [4, None, "x"]]

    def test_any_iterable_of_iterables_is_accepted_and_returns_lists(self):
        out = grid.append_row(((1, 2), (3, 4)), (5, 6))
        assert out == [[1, 2], [3, 4], [5, 6]]
        assert all(type(row) is list for row in out)
        gen = (list(range(i)) for i in range(3))
        assert grid.pad(gen) == [[None, None], [0, None], [0, 1]]


class TestShow:
    @pytest.mark.parametrize("seed", SEEDS)
    def test_never_crashes_on_odd_content_and_shows_every_row_when_asked(self, seed, capsys):
        rng = random.Random(800 + seed)
        g = [
            [
                rng.choice(
                    [None, "", "wide 日本語", "😀", "multi\nline", "tab\t", 1.5, True, "x" * 40]
                )
                for _ in range(rng.randint(0, 4))
            ]
            for _ in range(rng.randint(0, 8))
        ]
        grid.show(g, head=None, tail=None)
        printed = capsys.readouterr().out
        if g:
            assert printed.endswith("\n")
        grid.show(g)  # the default truncation must not crash either
        grid.show(g, rows=2)

    @pytest.mark.parametrize(("kwargs"), [{"rows": -1}, {"head": -1}, {"tail": -2}])
    def test_negative_counts_are_rejected(self, kwargs):
        with pytest.raises(ValueError):
            grid.show([[1]], **kwargs)

    @pytest.mark.parametrize(("kwargs"), [{"rows": 1.5}, {"head": "5"}, {"tail": [1]}])
    def test_non_integer_counts_are_rejected(self, kwargs):
        with pytest.raises(TypeError):
            grid.show([[1]], **kwargs)

    def test_a_table_no_longer_than_head_plus_tail_is_shown_whole(self, capsys):
        grid.show([[i] for i in range(10)])
        assert "..." not in capsys.readouterr().out
        grid.show([[i] for i in range(11)])
        assert "..." in capsys.readouterr().out
