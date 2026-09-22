"""Tests for Table column-type restrictions: construction, mutation,
enforcement on write/create, persistence, and equality."""

from __future__ import annotations

import datetime as dt

import pytest

from pyhandlexl import ColumnType, ColumnTypeError, Table, create_sheet


@pytest.fixture
def data_sheet(book):
    create_sheet(book, "Data")
    return book


class TestConstruction:
    def test_default_is_any_for_every_column(self):
        t = Table(data=[[1, "x"]], column_headers=["a", "b"], row_labels=["r"], name="T")
        assert t.column_types == {"a": ColumnType.ANY, "b": ColumnType.ANY}

    def test_declared_types_are_set(self):
        t = Table(
            data=[[1, "x"]],
            column_headers=["a", "b"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        assert t.column_types == {"a": ColumnType.NUMBER, "b": ColumnType.ANY}

    def test_unknown_header_raises_keyerror(self):
        with pytest.raises(KeyError):
            Table(
                data=[[1]],
                column_headers=["a"],
                row_labels=["r"],
                name="T",
                column_types={"ghost": ColumnType.NUMBER},
            )

    def test_non_column_type_value_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(
                data=[[1]],
                column_headers=["a"],
                row_labels=["r"],
                name="T",
                column_types={"a": "number"},
            )

    def test_non_mapping_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[[1]], column_headers=["a"], row_labels=["r"], name="T", column_types=["a"])

    def test_violating_data_is_rejected_immediately(self):
        with pytest.raises(ColumnTypeError):
            Table(
                data=[[1], ["not a number"]],
                column_headers=["a"],
                row_labels=["x", "y"],
                name="T",
                column_types={"a": ColumnType.NUMBER},
            )

    def test_bool_rejected_by_number(self):
        with pytest.raises(ColumnTypeError):
            Table(
                data=[[True]],
                column_headers=["a"],
                row_labels=["x"],
                name="T",
                column_types={"a": ColumnType.NUMBER},
            )


class TestSetColumnType:
    def test_changes_one_column(self):
        t = Table(data=[[1, "x"]], column_headers=["a", "b"], row_labels=["r"], name="T")
        t.set_column_type("b", ColumnType.TEXT)
        assert t.column_types == {"a": ColumnType.ANY, "b": ColumnType.TEXT}

    def test_unknown_header_raises_keyerror(self):
        t = Table(data=[[1]], column_headers=["a"], row_labels=["r"], name="T")
        with pytest.raises(KeyError):
            t.set_column_type("ghost", ColumnType.NUMBER)

    def test_non_column_type_raises_typeerror(self):
        t = Table(data=[[1]], column_headers=["a"], row_labels=["r"], name="T")
        with pytest.raises(TypeError):
            t.set_column_type("a", "number")

    def test_any_removes_restriction(self):
        t = Table(
            data=[[1]],
            column_headers=["a"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        t.set_column_type("a", ColumnType.ANY)
        assert t.column_types == {"a": ColumnType.ANY}


class TestEnforcementOnCreate:
    def test_matching_value_succeeds(self, data_sheet):
        t = Table(
            data=[[1], [2]],
            column_headers=["a"],
            row_labels=["x", "y"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        t.create(data_sheet, sheet="Data")
        assert Table.read(data_sheet, "T").data.rows == [[1], [2]]

    def test_none_is_always_allowed(self, data_sheet):
        t = Table(
            data=[[1], [None]],
            column_headers=["a"],
            row_labels=["x", "y"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        t.create(data_sheet, sheet="Data")
        assert Table.read(data_sheet, "T").data.rows == [[1], [None]]

    def test_date_and_datetime_both_allowed_by_date_type(self, data_sheet):
        t = Table(
            data=[[dt.date(2026, 1, 1)], [dt.datetime(2026, 1, 1, 9, 0)]],
            column_headers=["a"],
            row_labels=["x", "y"],
            name="T",
            column_types={"a": ColumnType.DATE},
        )
        t.create(data_sheet, sheet="Data")


class TestEnforcementOnEdit:
    """A value is checked against its column's restriction the moment it is set — not
    deferred to the next write()/create(). The one case still caught later is a
    restriction that arrives from another writer through a merge (see
    test_table_concurrency.py)."""

    @pytest.fixture
    def restricted(self) -> Table:
        return Table(
            data=[[1], [2]],
            column_headers=["a"],
            row_labels=["x", "y"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )

    def test_set_cell_by_label(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.set_cell(row="x", column="a", value="nope")
        assert restricted.read_cell(row="x", column="a") == 1

    def test_set_cell_by_position(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.set_cell(row=2, column=2, value="nope")
        assert restricted.read_cell(row=2, column=2) == 1

    def test_set_row(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.set_row("x", ["nope"])
        assert restricted.read_row("x") == [1]

    def test_set_column(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.set_column("a", [1, "nope"])
        assert restricted.read_column("a") == [1, 2]

    def test_add_row(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.add_row("z", ["nope"])
        assert "z" not in restricted.row_labels

    def test_insert_row(self, restricted):
        with pytest.raises(ColumnTypeError):
            restricted.insert_row(1, "z", ["nope"])
        assert "z" not in restricted.row_labels

    def test_set_column_type_against_existing_violating_data(self, data_sheet):
        t = Table(data=[["not a number"]], column_headers=["a"], row_labels=["x"], name="T")
        t.create(data_sheet, sheet="Data")  # fine — column is ANY at this point
        with pytest.raises(ColumnTypeError):
            t.set_column_type("a", ColumnType.NUMBER)
        assert t.column_types == {"a": ColumnType.ANY}  # rejected: nothing changed


class TestPersistence:
    def test_column_types_survive_read(self, data_sheet):
        Table(
            data=[[1, "x"]],
            column_headers=["a", "b"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER, "b": ColumnType.TEXT},
        ).create(data_sheet, sheet="Data")
        t = Table.read(data_sheet, "T")
        assert t.column_types == {"a": ColumnType.NUMBER, "b": ColumnType.TEXT}

    def test_change_via_set_column_type_persists_after_write(self, data_sheet):
        Table(data=[[1]], column_headers=["a"], row_labels=["r"], name="T").create(
            data_sheet, sheet="Data"
        )
        t = Table.read(data_sheet, "T")
        t.set_column_type("a", ColumnType.NUMBER)
        t.write(data_sheet)
        assert Table.read(data_sheet, "T").column_types == {"a": ColumnType.NUMBER}


class TestColumnShapeChanges:
    def test_insert_column_defaults_to_any(self):
        t = Table(
            data=[[1, 2]],
            column_headers=["a", "c"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER, "c": ColumnType.NUMBER},
        )
        t.insert_column(2, "b", [99])
        assert t.column_types == {
            "a": ColumnType.NUMBER,
            "b": ColumnType.ANY,
            "c": ColumnType.NUMBER,
        }

    def test_insert_in_the_middle_keeps_types_aligned(self):
        t = Table(
            data=[[1, "x"]],
            column_headers=["num", "txt"],
            row_labels=["r"],
            name="T",
            column_types={"num": ColumnType.NUMBER, "txt": ColumnType.TEXT},
        )
        t.insert_column(1, "new", [0])
        assert t.column_types == {
            "new": ColumnType.ANY,
            "num": ColumnType.NUMBER,
            "txt": ColumnType.TEXT,
        }

    def test_drop_column_removes_its_type(self):
        t = Table(
            data=[[1, "x"]],
            column_headers=["a", "b"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER, "b": ColumnType.TEXT},
        )
        t.drop_column("a")
        assert t.column_types == {"b": ColumnType.TEXT}

    def test_rename_column_carries_its_type_along(self):
        t = Table(
            data=[[1]],
            column_headers=["a"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        t.rename_column("a", "renamed")
        assert t.column_types == {"renamed": ColumnType.NUMBER}


class TestEquality:
    def test_different_column_types_are_not_equal(self):
        t1 = Table(data=[[1]], column_headers=["a"], row_labels=["r"], name="T")
        t2 = Table(
            data=[[1]],
            column_headers=["a"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        assert t1 != t2

    def test_same_column_types_are_equal(self):
        t1 = Table(
            data=[[1]],
            column_headers=["a"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        t2 = Table(
            data=[[1]],
            column_headers=["a"],
            row_labels=["r"],
            name="T",
            column_types={"a": ColumnType.NUMBER},
        )
        assert t1 == t2


class TestFromDict:
    def test_column_types_applied(self):
        t = Table.from_dict(
            {"r": {"a": 1, "b": "x"}}, name="T", column_types={"a": ColumnType.NUMBER}
        )
        assert t.column_types == {"a": ColumnType.NUMBER, "b": ColumnType.ANY}

    def test_unknown_header_raises_keyerror(self):
        with pytest.raises(KeyError):
            Table.from_dict({"r": {"a": 1}}, name="T", column_types={"ghost": ColumnType.NUMBER})
