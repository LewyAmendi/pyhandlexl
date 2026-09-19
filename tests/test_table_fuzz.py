"""Randomised tests against a plain-Python reference model.

A ``Table`` is driven through long random sequences of edits while a dead-simple model
(lists of lists) is updated alongside. After every step the table must agree with the
model, and at intervals it's written and read back from a real workbook. Nothing here
knows how ``Table`` works inside — a disagreement is a bug in one or the other.

Every test is seeded and parametrised, so a failure names the seed that reproduces it.
"""

from __future__ import annotations

import datetime as dt
import random
import warnings

import pytest

from pyhandlexl import MergeConflictWarning, Table, create_sheet, list_tables
from pyhandlexl._merge import same_value

# Values that survive an Excel round trip unchanged.
VALUES = [
    None,
    0,
    1,
    -7,
    123456789,
    3.5,
    -0.25,
    1e-9,
    "x",
    "text with spaces",
    "  padded  ",
    "007",
    "1e3",
    "TRUE",
    "é ü ñ",
    "日本語",
    "😀",
    "line1\nline2",
    "a,b;c\"d'e",
    True,
    False,
    dt.datetime(2026, 5, 17, 8, 30, 15),
    dt.datetime(1999, 12, 31, 23, 59, 59, 500000),
    dt.time(1, 2, 3),
    dt.timedelta(hours=5, minutes=6),
    dt.timedelta(days=-2),
]
LABELS = [
    "",
    "TRUE",
    "007",
    "1.5",
    "é",
    "日本",
    "a b",
    " lead",
    "trail ",
    "😀",
    "x/y",
    "q'r",
    "=notformula",
]

SEEDS = range(8)


def _same(a, b) -> bool:
    return same_value(a, b)


def _same_rows(left, right) -> bool:
    return len(left) == len(right) and all(
        len(x) == len(y) and all(_same(p, q) for p, q in zip(x, y, strict=True))
        for x, y in zip(left, right, strict=True)
    )


class Model:
    """The reference: labels, headers, and a list of value rows. No cleverness."""

    def __init__(self, labels, headers, rows, corner=""):
        self.labels = list(labels)
        self.headers = list(headers)
        self.rows = [list(r) for r in rows]
        self.corner = corner

    def check(self, table: Table, where: str) -> None:
        data = table.data
        assert data.row_labels == self.labels, where
        assert data.column_headers == self.headers, where
        assert data.corner == self.corner, where
        assert _same_rows(data.rows, self.rows), f"{where}: {data.rows} != {self.rows}"
        assert _same_rows(data.columns, [list(c) for c in zip(*self.rows, strict=True)])


class Driver:
    """Applies one random edit to a Table and to its Model, identically."""

    def __init__(self, table: Table, model: Model, rng: random.Random, prefix: str = "n") -> None:
        self.t, self.m, self.rng, self.prefix = table, model, rng, prefix
        self.counter = 0

    def fresh(self, pool) -> str:
        self.counter += 1
        base = self.rng.choice(pool)
        return f"{self.prefix}{self.counter}{base}"

    def value(self):
        return self.rng.choice(VALUES)

    def step(self) -> str:
        t, m, r = self.t, self.m, self.rng
        ops = ["add_row", "insert_row", "set_cell", "set_row", "set_corner", "add_column"]
        if m.labels:
            ops += ["drop_row", "rename_row", "set_column"] * 2 + ["set_cell"] * 3
        if len(m.headers) > 1:
            ops += ["drop_column", "rename_column"]
        ops += ["insert_column"]
        op = r.choice(ops)

        if op == "add_row":
            label, row = self.fresh(LABELS), [self.value() for _ in m.headers]
            t.add_row(label, row)
            m.labels.append(label)
            m.rows.append(row)
        elif op == "insert_row":
            pos = r.randint(1, len(m.labels) + 1)
            label, row = self.fresh(LABELS), [self.value() for _ in m.headers]
            t.insert_row(pos, label, row)
            m.labels.insert(pos - 1, label)
            m.rows.insert(pos - 1, row)
        elif op == "drop_row":
            i = r.randrange(len(m.labels))
            t.drop_row(m.labels[i])
            del m.labels[i], m.rows[i]
        elif op == "rename_row":
            i = r.randrange(len(m.labels))
            new = self.fresh(LABELS)
            t.rename_row(m.labels[i], new)
            m.labels[i] = new
        elif op == "set_cell":
            if not m.labels:
                return "noop"
            i, j = r.randrange(len(m.labels)), r.randrange(len(m.headers))
            v = self.value()
            t.set_cell(row=m.labels[i], column=m.headers[j], value=v)
            m.rows[i][j] = v
        elif op == "set_row":
            if not m.labels:
                return "noop"
            i = r.randrange(len(m.labels))
            row = [self.value() for _ in m.headers]
            t.set_row(m.labels[i], row)
            m.rows[i] = row
        elif op == "set_column":
            j = r.randrange(len(m.headers))
            col = [self.value() for _ in m.labels]
            t.set_column(m.headers[j], col)
            for row, v in zip(m.rows, col, strict=True):
                row[j] = v
        elif op == "add_column":
            header, col = self.fresh(LABELS), [self.value() for _ in m.labels]
            t.add_column(header, col)
            m.headers.append(header)
            for row, v in zip(m.rows, col, strict=True):
                row.append(v)
        elif op == "insert_column":
            pos = r.randint(1, len(m.headers) + 1)
            header, col = self.fresh(LABELS), [self.value() for _ in m.labels]
            t.insert_column(pos, header, col)
            m.headers.insert(pos - 1, header)
            for row, v in zip(m.rows, col, strict=True):
                row.insert(pos - 1, v)
        elif op == "drop_column":
            j = r.randrange(len(m.headers))
            t.drop_column(m.headers[j])
            del m.headers[j]
            for row in m.rows:
                del row[j]
        elif op == "rename_column":
            j = r.randrange(len(m.headers))
            new = self.fresh(LABELS)
            t.rename_column(m.headers[j], new)
            m.headers[j] = new
        elif op == "set_corner":
            m.corner = r.choice(["", "corner", "é", "😀", "Total", "a\nb"])
            t.set_corner(m.corner)
        return op


def _build(rng: random.Random, name: str, prefix: str):
    n_cols = rng.randint(1, 4)
    headers = [f"{prefix}h{j}{rng.choice(LABELS)}" for j in range(n_cols)]
    n_rows = rng.randint(0, 4)
    labels = [f"{prefix}r{i}{rng.choice(LABELS)}" for i in range(n_rows)]
    rows = [[rng.choice(VALUES) for _ in headers] for _ in labels]
    table = Table(rows, labels, column_headers=headers, name=name)
    return table, Model(labels, headers, rows)


@pytest.fixture
def path(book):
    create_sheet(book, "Data")
    create_sheet(book, "More")
    return book


class TestOneTableAgainstTheModel:
    @pytest.mark.parametrize("seed", SEEDS)
    def test_every_step_and_every_round_trip_agrees(self, path, seed):
        rng = random.Random(seed)
        table, model = _build(rng, "T", "")
        table.create(path, sheet="Data")
        driver = Driver(table, model, rng)
        for step in range(24):
            op = driver.step()
            model.check(table, f"seed {seed} step {step} after {op}")
            if rng.random() < 0.25:
                table.write(path)
                model.check(table, f"seed {seed} step {step} after write")
                model.check(Table.read(path, "T"), f"seed {seed} step {step} read back")
        table.write(path)
        model.check(Table.read(path, "T"), f"seed {seed} final")
        assert list_tables(path) == ["T"]


class TestSeveralTablesShareASheet:
    @pytest.mark.parametrize("seed", SEEDS)
    def test_growing_and_shrinking_one_table_never_disturbs_the_others(self, path, seed):
        rng = random.Random(1000 + seed)
        names = ["A", "B", "C"]
        drivers = {}
        for i, name in enumerate(names):
            table, model = _build(rng, name, f"{name.lower()}_")
            table.create(path, sheet="Data" if i < 2 else "More")
            drivers[name] = Driver(table, model, rng, prefix=name.lower())
        # C shares a sheet with nobody; A and B share "Data" side by side.
        for step in range(8):
            name = rng.choice(names)
            drivers[name].step()
            drivers[name].t.write(path)
            for other, drv in drivers.items():
                drv.m.check(Table.read(path, other), f"seed {seed} step {step} table {other}")
        assert sorted(list_tables(path)) == names


def _random_value_edit(rng, table: Table, tag: str) -> None:
    """A harmless-to-merge random edit: touches one existing cell or adds a tagged row/column."""
    labels, headers = table.data.row_labels, table.data.column_headers
    kind = rng.choice(["cell", "cell", "row", "column"])
    if kind == "cell" and labels:
        table.set_cell(
            row=rng.choice(labels), column=rng.choice(headers), value=f"{tag}{rng.randint(0, 99)}"
        )
    elif kind == "row":
        table.add_row(f"{tag}row{rng.randint(0, 10**6)}", [f"{tag}v"] * len(headers))
    else:
        table.add_column(f"{tag}col{rng.randint(0, 10**6)}", [f"{tag}c"] * len(labels))


def _make_base(path, rng, name="S"):
    n_rows, n_cols = rng.randint(2, 6), rng.randint(2, 4)
    labels = [f"r{i}" for i in range(n_rows)]
    headers = [f"c{j}" for j in range(n_cols)]
    rows = [[f"{i}:{j}" for j in range(n_cols)] for i in range(n_rows)]
    Table(rows, labels, column_headers=headers, name=name).create(path, sheet="Data")
    return labels, headers, rows


class TestTwoWritersWithNothingInCommon:
    """Alice and Bob touch different rows and add different things: nothing may be lost."""

    @pytest.mark.parametrize("seed", SEEDS)
    def test_the_merge_is_exactly_the_union_of_both(self, path, seed):
        rng = random.Random(2000 + seed)
        labels, headers, rows = _make_base(path, rng)
        alice, bob = Table.read(path, "S"), Table.read(path, "S")
        mine = {"alice": set(), "bob": set()}
        # give each writer their own base rows to edit, and their own additions
        owners = {label: rng.choice(["alice", "bob"]) for label in labels}
        expected = {
            (labels[i], headers[j]): rows[i][j]
            for i in range(len(labels))
            for j in range(len(headers))
        }
        added_rows, added_cols = {"alice": [], "bob": []}, {"alice": [], "bob": []}
        for who, table in (("alice", alice), ("bob", bob)):
            for _ in range(rng.randint(0, 6)):
                kind = rng.choice(["cell", "row", "column"])
                if kind == "cell":
                    owned = [x for x in labels if owners[x] == who]
                    if owned:
                        label, header = rng.choice(owned), rng.choice(headers)
                        value = f"{who}-{rng.randint(0, 999)}"
                        table.set_cell(row=label, column=header, value=value)
                        expected[(label, header)] = value
                        mine[who].add((label, header))
                elif kind == "row":
                    label = f"{who}-row{len(added_rows[who])}"
                    table.add_row(label, [f"{who}-new"] * len(table.data.column_headers))
                    added_rows[who].append(label)
                else:
                    header = f"{who}-col{len(added_cols[who])}"
                    table.add_column(header, [f"{who}-col"] * len(table.data.row_labels))
                    added_cols[who].append(header)
        with warnings.catch_warnings():
            warnings.simplefilter("error", MergeConflictWarning)  # nothing collides
            alice.write(path)
            bob.write(path)
        final = Table.read(path, "S")
        assert set(final.data.row_labels) == set(labels) | set(added_rows["alice"]) | set(
            added_rows["bob"]
        )
        assert set(final.data.column_headers) == set(headers) | set(added_cols["alice"]) | set(
            added_cols["bob"]
        )
        assert len(set(final.data.row_labels)) == len(final.data.row_labels)
        # every base cell holds whoever edited it (or the original)
        for (label, header), value in expected.items():
            assert final.read_cell(row=label, column=header) == value, (label, header)
        # a row one writer added has that writer's value in the columns they knew about
        for who in ("alice", "bob"):
            for label in added_rows[who]:
                for header in headers:
                    assert final.read_cell(row=label, column=header) == f"{who}-new"
            for header in added_cols[who]:
                for label in labels:
                    assert final.read_cell(row=label, column=header) == f"{who}-col"
        # the two additions never leave a stray value in the corner they both didn't know
        for label in added_rows["alice"]:
            for header in added_cols["bob"]:
                assert final.read_cell(row=label, column=header) is None
        for label in added_rows["bob"]:
            for header in added_cols["alice"]:
                assert final.read_cell(row=label, column=header) is None
        assert bob.data == final.data
        assert bob.last_merge is None or not bob.last_merge.conflicts


class TestTwoWritersDoingAnything:
    """Random edits by both, conflicts and all: the table must stay valid and self-consistent."""

    @pytest.mark.parametrize("seed", SEEDS)
    def test_the_result_is_always_a_valid_table_equal_to_what_the_second_writer_holds(
        self, path, seed
    ):
        rng = random.Random(3000 + seed)
        _make_base(path, rng)
        alice, bob = Table.read(path, "S"), Table.read(path, "S")
        for who, table in (("alice", alice), ("bob", bob)):
            model = Model(
                table.data.row_labels, table.data.column_headers, table.data.rows, table.data.corner
            )
            driver = Driver(table, model, rng, prefix=who[0])
            for _ in range(rng.randint(0, 8)):
                driver.step()
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", MergeConflictWarning)
            alice.write(path)
            bob.write(path)
        final = Table.read(path, "S")
        labels, headers = final.data.row_labels, final.data.column_headers
        assert len(set(labels)) == len(labels) and len(set(headers)) == len(headers)
        assert headers, "a table always keeps at least one column"
        assert all(len(row) == len(headers) for row in final.data.rows)
        assert len(final.data.rows) == len(labels)
        assert _same_rows(final.data.rows, bob.data.rows)
        assert final.data.row_labels == bob.data.row_labels
        assert final.data.column_headers == bob.data.column_headers
        # a write with nothing new on disk is a no-op and doesn't disturb anything
        bob.write(path)
        assert bob.last_merge is None
        again = Table.read(path, "S")
        assert again.data.row_labels == labels and _same_rows(again.data.rows, final.data.rows)

    @pytest.mark.parametrize("seed", SEEDS)
    def test_a_writer_who_changed_nothing_never_disturbs_the_other(self, path, seed):
        rng = random.Random(4000 + seed)
        _make_base(path, rng)
        alice, bob = Table.read(path, "S"), Table.read(path, "S")
        for _ in range(rng.randint(1, 6)):
            _random_value_edit(rng, alice, "a")
        alice.write(path)
        after_alice = Table.read(path, "S")
        bob.write(path)  # bob has changed nothing
        final = Table.read(path, "S")
        assert final.data.row_labels == after_alice.data.row_labels
        assert final.data.column_headers == after_alice.data.column_headers
        assert _same_rows(final.data.rows, after_alice.data.rows)
        assert bob.last_merge is not None and not bob.last_merge.conflicts
