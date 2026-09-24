"""Wrong arguments get a deliberate error — never an accident.

Every public function and every ``Table`` method is called with garbage in each parameter in turn
(``None``, numbers, bytes, containers, ``object()`` …). Any of the standard argument errors, a file
error, or one of the library's own is fine; anything else — an ``AttributeError`` from calling
``.strip()`` on a tuple, a ``RuntimeError`` — is a bug in the API's argument handling.
"""

from __future__ import annotations

import contextlib
import datetime as dt
import inspect
import io
import itertools
import shutil

import pytest

import pyhandlexl as p
from pyhandlexl import grid

DELIBERATE = (
    TypeError,
    ValueError,
    KeyError,
    IndexError,
    FileNotFoundError,
    FileExistsError,
    IsADirectoryError,
    NotADirectoryError,
    ImportError,
    p.PyhandlexlError,
)

GARBAGE = [
    None,
    5,
    2.5,
    True,
    b"bytes",
    [],
    {},
    object(),
    ("t",),
    {"a"},
    "",
    "a" * 40,
    dt.date(2026, 1, 1),
    float("nan"),
    p.TableStyle.DEFAULT,
]

# a sensible value for each parameter name, so that varying one parameter at a time reaches
# past the earliest checks
GOOD = {
    "sheet": "Sheet",
    "name": "T",
    "old": "a",
    "new": "b",
    "label": "a",
    "header": "x",
    "position": 1,
    "values": [1, 2],
    "value": 1,
    "column_type": p.ColumnType.ANY,
    "ref": "B2",
    "row": "a",
    "column": "x",
    "rows": [[1, 2]],
    "head": 5,
    "tail": 5,
    "orientation": "rows",
    "pad": False,
    "encoding": "utf-8",
    "n_rows": 3,
    "n_cols": 3,
    "grid": [[1, 2], [3, 4]],
    "row_labels": ["a"],
    "column_headers": ["x", "y"],
    "corner": "",
    "style": None,
    "column_types": None,
    "table": {"a": {"x": 1}},
    "infer_column_types": False,
    "index": False,
    "dtype": None,
    "data": [[1, 2]],
    "df": None,
    "col": 1,
}


@pytest.fixture(scope="module")
def base(tmp_path_factory):
    path = tmp_path_factory.mktemp("misuse") / "base.xlsx"
    p.create_workbook(path)
    p.create_sheet(path, "D")
    p.write_sheet(path, [[1, 2], [3, 4]], sheet="Sheet")
    p.Table(data=[[1, 2]], row_labels=["a"], column_headers=["x", "y"], name="T").create(
        path, sheet="D"
    )
    return path


@pytest.fixture
def fresh(base, tmp_path):
    counter = itertools.count()

    def make():
        copy = tmp_path / f"w{next(counter)}.xlsx"
        shutil.copyfile(base, copy)
        return str(copy)

    return make


def table() -> p.Table:
    return p.Table(data=[[1, 2]], row_labels=["a"], column_headers=["x", "y"], name="T")


def arguments(func, replace: str, garbage, fresh):
    """The call arguments for *func*, everything sensible except parameter *replace*."""
    parameters = [
        q
        for q in inspect.signature(func).parameters.values()
        if q.name not in ("self", "cls")
        and q.kind in (q.POSITIONAL_ONLY, q.POSITIONAL_OR_KEYWORD, q.KEYWORD_ONLY)
    ]
    positional, keyword = [], {}
    for q in parameters:
        if q.name == replace:
            value = garbage
        elif q.name in ("path", "csv_path"):
            value = fresh()
        else:
            value = GOOD.get(q.name, "x")
        if q.kind is q.POSITIONAL_ONLY:
            positional.append(value)
        else:
            keyword[q.name] = value
    return parameters, positional, keyword


def check_every_parameter(func, fresh, make_call):
    accidents = []
    parameters, _, _ = arguments(func, "", None, fresh)
    for q in parameters:
        for garbage in GARBAGE:
            _, positional, keyword = arguments(func, q.name, garbage, fresh)
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    make_call(positional, keyword)
            except DELIBERATE:
                continue
            except BaseException as error:
                accidents.append(
                    f"{q.name}={garbage!r:.30} -> {type(error).__name__}: {error!s:.60}"
                )
    assert not accidents, "accidental exceptions:\n  " + "\n  ".join(accidents[:8])


FUNCTIONS = [
    name
    for name in p.__all__
    if inspect.isfunction(getattr(p, name)) and name not in {"delete_workbook", "rename_workbook"}
]
GRID_FUNCTIONS = [
    name
    for name, member in inspect.getmembers(grid, inspect.isfunction)
    if not name.startswith("_") and member.__module__ == grid.__name__
]
TABLE_METHODS = [
    name
    for name, member in vars(p.Table).items()
    if (not name.startswith("_") or name == "__init__")
    and not isinstance(member, property)
    and inspect.isfunction(member.__func__ if hasattr(member, "__func__") else member)
]


@pytest.mark.parametrize("name", FUNCTIONS)
def test_a_public_function(name, fresh):
    func = getattr(p, name)
    check_every_parameter(func, fresh, lambda positional, keyword: func(*positional, **keyword))


@pytest.mark.parametrize("name", GRID_FUNCTIONS)
def test_a_grid_function(name, fresh):
    func = getattr(grid, name)
    check_every_parameter(func, fresh, lambda positional, keyword: func(*positional, **keyword))


@pytest.mark.parametrize("name", TABLE_METHODS)
def test_a_table_method(name, fresh):
    member = vars(p.Table)[name]
    func = member.__func__ if hasattr(member, "__func__") else member

    def call(positional, keyword):
        if name == "__init__":
            return p.Table(*positional, **{"column_headers": ["x"], "name": "T", **keyword})
        if isinstance(member, classmethod):
            return getattr(p.Table, name)(*positional, **keyword)
        return getattr(table(), name)(*positional, **keyword)

    check_every_parameter(func, fresh, call)


class TestASheetNameThatIsNotAString:
    """The four functions that used to fail on ``.strip()``, and the rest for good measure."""

    @pytest.mark.parametrize("bad", [5, 2.5, b"x", ("t",), ["a"], object()], ids=repr)
    def test_every_function_that_takes_one_refuses_it_deliberately(self, bad, fresh):
        calls = [
            lambda: p.Table(column_headers=["a"], name="N").create(fresh(), sheet=bad),
            lambda: p.clear_all_sheet_data(fresh(), bad),
            lambda: p.delete_sheet(fresh(), bad),
            lambda: p.sheet_kind(fresh(), bad),
            lambda: p.create_sheet(fresh(), bad),
            lambda: p.rename_sheet(fresh(), bad, "Other"),
            lambda: p.rename_sheet(fresh(), "Sheet", bad),
            lambda: p.write_sheet(fresh(), [[1]], sheet=bad),
            lambda: p.append_rows(fresh(), [[1]], sheet=bad),
            lambda: p.read_sheet(fresh(), bad),
        ]
        for call in calls:
            with pytest.raises(p.PyhandlexlError):
                call()

    def test_none_is_refused_where_a_name_is_required(self, fresh):
        # (write_sheet, append_rows and read_sheet take None to mean "the active sheet")
        calls = [
            lambda: p.Table(column_headers=["a"], name="N").create(fresh(), sheet=None),
            lambda: p.clear_all_sheet_data(fresh(), None),
            lambda: p.delete_sheet(fresh(), None),
            lambda: p.sheet_kind(fresh(), None),
            lambda: p.create_sheet(fresh(), None),
            lambda: p.rename_sheet(fresh(), None, "Other"),
        ]
        for call in calls:
            with pytest.raises(p.SheetNameError):
                call()

    def test_the_ones_that_used_to_crash_say_why(self, fresh):
        with pytest.raises(p.SheetNameError, match="must be a string, got tuple"):
            p.sheet_kind(fresh(), ("t",))
        with pytest.raises(p.SheetNameError, match="must be a string, got int"):
            p.delete_sheet(fresh(), 5)
