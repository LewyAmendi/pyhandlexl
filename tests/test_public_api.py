"""The public API is a promise: this pins its exact surface.

Every name in ``pyhandlexl.__all__``, every public method and property of its classes, the
fields of its value types, its exceptions' bases and the functions of ``pyhandlexl.grid``,
with their signatures. If a test here fails, the public API changed. That is fine when it is
deliberate — update the table below *and* say so in the changelog — and a bug when it isn't.
(To regenerate the table after a deliberate change, print ``describe()``.)
"""

from __future__ import annotations

import dataclasses
import difflib
import enum
import inspect

import pyhandlexl
from pyhandlexl import grid


def _sig(obj: object) -> str:
    return str(inspect.signature(obj))  # type: ignore[arg-type]


def _describe_class(cls: type) -> list[str]:
    lines = [f"class {cls.__name__}({', '.join(b.__name__ for b in cls.__bases__)})"]
    if issubclass(cls, enum.Enum):
        lines += [f"  {cls.__name__}.{m.name} = {m.value!r}" for m in cls]
    if dataclasses.is_dataclass(cls):
        frozen = cls.__dataclass_params__.frozen  # type: ignore[attr-defined]
        lines.append(f"  dataclass(frozen={frozen})")
        lines += [f"  field {f.name}: {f.type}" for f in dataclasses.fields(cls)]
    for name, member in sorted(vars(cls).items()):
        if name.startswith("_") and name != "__init__":
            continue
        if isinstance(member, property):
            lines.append(f"  property {name}{' (settable)' if member.fset else ''}")
        elif isinstance(member, classmethod | staticmethod):
            kind = type(member).__name__
            lines.append(f"  {kind} {name}{_sig(member.__func__)}")
        elif inspect.isfunction(member):
            lines.append(f"  def {name}{_sig(member)}")
    return lines


def describe() -> list[str]:
    lines: list[str] = []
    for name in sorted(pyhandlexl.__all__):
        obj = getattr(pyhandlexl, name)
        if inspect.isclass(obj):
            lines += _describe_class(obj)
        elif inspect.isfunction(obj):
            lines.append(f"def {name}{_sig(obj)}")
        else:
            lines.append(f"{name}: {type(obj).__name__}")
    for name, obj in sorted(inspect.getmembers(grid, inspect.isfunction)):
        if not name.startswith("_") and obj.__module__ == grid.__name__:
            lines.append(f"def grid.{name}{_sig(obj)}")
    return lines


def test_the_public_api_is_unchanged():
    actual = describe()
    if actual != EXPECTED:
        diff = "\n".join(difflib.unified_diff(EXPECTED, actual, "expected", "actual", lineterm=""))
        raise AssertionError(f"the public API changed:\n{diff}")


def test_every_public_name_is_importable_and_listed_once():
    assert len(pyhandlexl.__all__) == len(set(pyhandlexl.__all__))
    for name in pyhandlexl.__all__:
        assert hasattr(pyhandlexl, name), name


def test_the_version_is_a_plain_release_number():
    assert pyhandlexl.__version__.count(".") == 2
    assert all(part.isdigit() for part in pyhandlexl.__version__.split("."))


EXPECTED = [
    "class CellTypeError(PyhandlexlError, TypeError)",
    "class ColumnType(Enum)",
    "  ColumnType.ANY = 'any'",
    "  ColumnType.TEXT = 'text'",
    "  ColumnType.NUMBER = 'number'",
    "  ColumnType.BOOLEAN = 'boolean'",
    "  ColumnType.DATE = 'date'",
    "  ColumnType.TIME = 'time'",
    "  ColumnType.DURATION = 'duration'",
    "  def allows(self, value: 'object') -> 'bool'",
    "class ColumnTypeError(PyhandlexlError, TypeError)",
    "class DimensionError(PyhandlexlError, ValueError)",
    "class FileLockedError(PyhandlexlError, OSError)",
    "class FileReadOnlyError(FileLockedError)",
    "class FormulaWarning(UserWarning)",
    "class InvalidFileError(PyhandlexlError, ValueError)",
    "class MalformedTableError(PyhandlexlError, ValueError)",
    "class MergeConflictWarning(UserWarning)",
    "class MergeReport(object)",
    "  dataclass(frozen=True)",
    "  field rows_added: tuple[str, ...]",
    "  field rows_removed: tuple[str, ...]",
    "  field columns_added: tuple[str, ...]",
    "  field columns_removed: tuple[str, ...]",
    "  field cells_updated: int",
    "  field conflicts: tuple[str, ...]",
    "  def __init__(self, rows_added: 'tuple[str, ...]' = (), rows_removed: 'tuple[str, ...]' = (), columns_added: 'tuple[str, ...]' = (), columns_removed: 'tuple[str, ...]' = (), cells_updated: 'int' = 0, conflicts: 'tuple[str, ...]' = ()) -> None",
    "Orientation: _LiteralGenericAlias",
    "class PyhandlexlError(Exception)",
    "class SchemaRebuiltWarning(UserWarning)",
    "class SheetExistsError(SheetNameError)",
    "SheetKind: _LiteralGenericAlias",
    "class SheetKindError(PyhandlexlError, ValueError)",
    "class SheetNameError(PyhandlexlError, ValueError)",
    "class SheetNotFoundError(PyhandlexlError, KeyError)",
    "class Table(object)",
    "  def __init__(self, data: 'Iterable[Iterable[object]]' = (), row_labels: 'Iterable[str]' = (), corner: 'str' = '', *, column_headers: 'Iterable[str]', name: 'str', style: 'TableStyle | None' = None, column_types: 'Mapping[str, ColumnType] | None' = None) -> 'None'",
    "  def add_column(self, header: 'str', values: 'Iterable[object]') -> 'None'",
    "  def add_row(self, label: 'str', values: 'Iterable[object]') -> 'None'",
    "  property column_headers",
    "  property column_types",
    "  property corner",
    "  def create(self, path: 'str | Path', sheet: 'str') -> 'None'",
    "  property data",
    "  def drop_column(self, header: 'str') -> 'None'",
    "  def drop_row(self, label: 'str') -> 'None'",
    "  classmethod from_dataframe(cls, df: 'Any', *, name: 'str', style: 'TableStyle | None' = None, column_types: 'Mapping[str, ColumnType] | None' = None, infer_column_types: 'bool' = False) -> 'Table'",
    "  classmethod from_dict(cls, table: 'Mapping[str, Mapping[str, object]]', corner: 'str' = '', *, name: 'str', style: 'TableStyle | None' = None, column_types: 'Mapping[str, ColumnType] | None' = None) -> 'Table'",
    "  property info",
    "  def insert_column(self, position: 'int', header: 'str', values: 'Iterable[object]') -> 'None'",
    "  def insert_row(self, position: 'int', label: 'str', values: 'Iterable[object]') -> 'None'",
    "  property last_merge",
    "  property name",
    "  classmethod read(cls, path: 'str | Path', name: 'str') -> 'Table'",
    "  def read_cell(self, ref: 'str | None' = None, *, row: 'int | str | None' = None, column: 'int | str | None' = None) -> 'object'",
    "  def read_column(self, header: 'str') -> 'list[object]'",
    "  def read_row(self, label: 'str') -> 'list[object]'",
    "  def rename_column(self, old: 'str', new: 'str') -> 'None'",
    "  def rename_row(self, old: 'str', new: 'str') -> 'None'",
    "  property row_labels",
    "  def set_cell(self, ref: 'str | None' = None, *, row: 'int | str | None' = None, column: 'int | str | None' = None, value: 'object') -> 'None'",
    "  def set_column(self, header: 'str', values: 'Iterable[object]') -> 'None'",
    "  def set_column_type(self, header: 'str', column_type: 'ColumnType') -> 'None'",
    "  def set_corner(self, value: 'str') -> 'None'",
    "  def set_row(self, label: 'str', values: 'Iterable[object]') -> 'None'",
    "  def show(self, *, rows: 'int | None' = None, head: 'int | None' = 5, tail: 'int | None' = 5) -> 'str'",
    "  property style (settable)",
    "  def to_dataframe(self) -> 'Any'",
    "  def to_dict(self) -> 'dict[str, dict[str, object]]'",
    "  def to_numpy(self, dtype: 'Any' = None) -> 'Any'",
    "  def write(self, path: 'str | Path') -> 'None'",
    "class TableData(object)",
    "  dataclass(frozen=True)",
    "  field rows: list[list[object]]",
    "  field columns: list[list[object]]",
    "  field row_labels: list[str]",
    "  field column_headers: list[str]",
    "  field corner: str",
    "  def __init__(self, rows: 'list[list[object]]', columns: 'list[list[object]]', row_labels: 'list[str]', column_headers: 'list[str]', corner: 'str') -> None",
    "class TableExistsError(PyhandlexlError, ValueError)",
    "class TableInfo(object)",
    "  dataclass(frozen=True)",
    "  field created_at: datetime | None",
    "  field modified_at: datetime | None",
    "  field n_rows: int",
    "  field n_cols: int",
    "  field sheet: str | None",
    "  field style: TableStyle",
    "  field column_types: dict[str, ColumnType]",
    "  def __init__(self, created_at: 'datetime | None', modified_at: 'datetime | None', n_rows: 'int', n_cols: 'int', sheet: 'str | None', style: 'TableStyle', column_types: 'dict[str, ColumnType]') -> None",
    "class TableNotFoundError(PyhandlexlError, KeyError)",
    "class TableStyle(object)",
    "  dataclass(frozen=True)",
    "  field header_font_name: str",
    "  field header_font_size: int",
    "  field header_font_color: str",
    "  field header_bold: bool",
    "  field header_fill: str",
    "  field data_font_name: str",
    "  field data_font_size: int",
    "  field data_font_color: str",
    "  field band_fill: str",
    "  field border_color: str",
    "  def __init__(self, header_font_name: 'str' = 'Calibri', header_font_size: 'int' = 11, header_font_color: 'str' = '000000', header_bold: 'bool' = True, header_fill: 'str' = '76933C', data_font_name: 'str' = 'Calibri', data_font_size: 'int' = 11, data_font_color: 'str' = '000000', band_fill: 'str' = 'F2F2F2', border_color: 'str' = '000000') -> None",
    "def append_rows(path: 'str | Path', rows: 'Iterable[Iterable[object]]', sheet: 'str | None' = None) -> 'None'",
    "def check_cell_value(value: 'object', /) -> 'None'",
    "def check_dimensions(n_rows: 'int', n_cols: 'int') -> 'None'",
    "def check_sheet_name(name: 'str', /) -> 'None'",
    "def clear_all_sheet_data(path: 'str | Path', sheet: 'str') -> 'None'",
    "def create_csv(path: 'str | Path') -> 'None'",
    "def create_sheet(path: 'str | Path', name: 'str') -> 'None'",
    "def create_workbook(path: 'str | Path', *, sheet: 'str' = 'Sheet') -> 'None'",
    "def delete_schema_sheet(path: 'str | Path') -> 'None'",
    "def delete_sheet(path: 'str | Path', name: 'str') -> 'None'",
    "def delete_table(path: 'str | Path', name: 'str') -> 'None'",
    "def delete_workbook(path: 'str | Path') -> 'None'",
    "def export_xl_to_csv(path: 'str | Path', csv_path: 'str | Path', *, sheet: 'str | None' = None, encoding: 'str' = 'utf-8') -> 'None'",
    "grid: module",
    "def import_csv_to_xl(csv_path: 'str | Path', path: 'str | Path', *, sheet: 'str | None' = None, encoding: 'str' = 'utf-8-sig') -> 'None'",
    "def is_valid_xlsx(path: 'str | Path', /) -> 'bool'",
    "def list_sheets(path: 'str | Path') -> 'list[str]'",
    "def list_tables(path: 'str | Path') -> 'list[str]'",
    "def read_sheet(path: 'str | Path', sheet: 'str | None' = None, *, pad: 'bool' = False, orientation: 'Orientation' = 'rows') -> 'list[list[object]]'",
    "def rename_sheet(path: 'str | Path', old: 'str', new: 'str') -> 'None'",
    "def rename_workbook(path: 'str | Path', new_name: 'str') -> 'None'",
    "def sheet_exists(path: 'str | Path', name: 'str') -> 'bool'",
    "def sheet_kind(path: 'str | Path', sheet: 'str') -> 'SheetKind'",
    "def table_info(path: 'str | Path', name: 'str') -> 'TableInfo'",
    "def write_sheet(path: 'str | Path', rows: 'Iterable[Iterable[object]]', sheet: 'str | None' = None, *, orientation: 'Orientation' = 'rows') -> 'None'",
    "def grid.append_column(grid: 'Iterable[Iterable[object]]', values: 'Iterable[object]') -> 'Grid'",
    "def grid.append_row(grid: 'Iterable[Iterable[object]]', values: 'Iterable[object]') -> 'Grid'",
    "def grid.delete_column(grid: 'Iterable[Iterable[object]]', col: 'int') -> 'Grid'",
    "def grid.delete_row(grid: 'Iterable[Iterable[object]]', row: 'int') -> 'Grid'",
    "def grid.from_dataframe(df: 'Any', *, header: 'bool' = True, index: 'bool' = False) -> 'Grid'",
    "def grid.get_column(grid: 'Iterable[Iterable[object]]', col: 'int') -> 'list[object]'",
    "def grid.get_row(grid: 'Iterable[Iterable[object]]', row: 'int') -> 'list[object]'",
    "def grid.insert_column(grid: 'Iterable[Iterable[object]]', col: 'int', values: 'Iterable[object]') -> 'Grid'",
    "def grid.insert_row(grid: 'Iterable[Iterable[object]]', row: 'int', values: 'Iterable[object]') -> 'Grid'",
    "def grid.pad(grid: 'Iterable[Iterable[object]]') -> 'Grid'",
    "def grid.set_column(grid: 'Iterable[Iterable[object]]', col: 'int', values: 'Iterable[object]') -> 'Grid'",
    "def grid.set_row(grid: 'Iterable[Iterable[object]]', row: 'int', values: 'Iterable[object]') -> 'Grid'",
    "def grid.set_value(grid: 'Iterable[Iterable[object]]', row: 'int', col: 'int', value: 'object') -> 'Grid'",
    "def grid.show(grid: 'Iterable[Iterable[object]]', *, rows: 'int | None' = None, head: 'int | None' = 5, tail: 'int | None' = 5) -> 'str'",
    "def grid.to_dataframe(grid: 'Iterable[Iterable[object]]', *, header: 'bool' = True) -> 'Any'",
    "def grid.transpose(grid: 'Iterable[Iterable[object]]') -> 'Grid'",
]
