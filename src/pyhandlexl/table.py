"""The Table class: a named table living on a worksheet.

Layout convention, anchored wherever the table's schema entry says it starts:
the first cell holds the marker ``"TABLE NAME"`` plus the table's name, the
next row holds the column headers (with the "corner" in its first cell), and
every row after that holds a row label followed by data. Row labels, column
headers, and the corner are always strings; data values keep their type.
Positional access uses coordinates relative to the table itself (row 1 is the
header row, column 1 is the label column).

Every table has a name — see :meth:`Table.create`, :meth:`Table.read`, and
:meth:`Table.write`. Multiple tables can share one worksheet: see
`Organised data: the Table class` in the README for the placement and
self-healing rules.
"""

from __future__ import annotations

import warnings
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path

from openpyxl.utils import coordinate_to_tuple

from pyhandlexl import _multi_table as mt
from pyhandlexl._merge import MergeReport, Snapshot, merge, same_state
from pyhandlexl._safety import atomic_save, safe_load
from pyhandlexl.column_type import ColumnType
from pyhandlexl.errors import (
    ColumnTypeError,
    MergeConflictWarning,
    SheetKindError,
    SheetNotFoundError,
)
from pyhandlexl.style import TableStyle
from pyhandlexl.validate import check_cell_value, check_dimensions


@dataclass(frozen=True)
class TableData:
    """A read-only snapshot of a Table's content."""

    rows: list[list[object]]
    columns: list[list[object]]
    row_labels: list[str]
    column_headers: list[str]
    corner: str


@dataclass(frozen=True)
class TableInfo:
    """A read-only snapshot of a Table's metadata — not its data.

    ``sheet``, ``created_at``, and ``modified_at`` are ``None`` until the
    table has been placed with :meth:`Table.create` (or loaded with
    :meth:`Table.read`) — a freshly built, not-yet-placed ``Table`` has no
    sheet or history yet. ``created_at``/``modified_at`` are also ``None``
    for a table recovered by an automatic schema rebuild (see
    :class:`~pyhandlexl.errors.SchemaRebuiltWarning`) — a rebuild has no way
    to know either time, so it reports them as unknown rather than guessing.

    ``modified_at`` only reflects this table's *own* ``create()``/``write()``
    calls — a table shifted right to make room for a growing neighbour (see
    "Multiple named tables on one sheet" in the README) is not itself
    modified, so its ``modified_at`` is unaffected by that.
    """

    created_at: datetime | None
    modified_at: datetime | None
    n_rows: int
    n_cols: int
    sheet: str | None
    style: TableStyle
    column_types: dict[str, ColumnType]


def _conflict_message(name: str, report: MergeReport) -> str:
    shown = report.conflicts[:10]
    lines = [f"  - {conflict}" for conflict in shown]
    if len(report.conflicts) > len(shown):
        lines.append(f"  ... and {len(report.conflicts) - len(shown)} more")
    return (
        f"table {name!r} was changed on disk since it was read; write() merged those changes "
        "with yours, and yours won where you both changed the same thing "
        f"({len(report.conflicts)}):\n" + "\n".join(lines)
    )


class Table:
    """A named table, always with column headers, row labels, and a corner.

    Construct directly from parts (``column_headers=`` and ``name=`` are
    required; a table with no rows yet still needs its columns defined), or
    with :meth:`read` from a file. Mutation methods edit the table in place
    and return ``None``. Row labels and column headers are always ``str``.

    Every table is visually styled — see :class:`pyhandlexl.style.TableStyle`
    — defaulting to ``TableStyle.DEFAULT`` unless ``style=`` says otherwise.

    A column can be restricted to one :class:`~pyhandlexl.column_type.ColumnType`
    — every column defaults to ``ColumnType.ANY`` (no restriction) unless
    ``column_types=`` says otherwise.
    """

    def __init__(
        self,
        data: Iterable[Iterable[object]] = (),
        row_labels: Iterable[str] = (),
        corner: str = "",
        *,
        column_headers: Iterable[str],
        name: str,
        style: TableStyle | None = None,
        column_types: Mapping[str, ColumnType] | None = None,
    ) -> None:
        # A bare str is technically Iterable[str] — iterating it silently splits
        # it into one column/row per character. That's never what's meant, so
        # it's rejected outright rather than left as a silent footgun.
        for value, what in (
            (data, "data"),
            (column_headers, "column_headers"),
            (row_labels, "row_labels"),
        ):
            if isinstance(value, (str, bytes)):
                raise TypeError(
                    f"{what} must be a sequence of items, not a single "
                    f"{type(value).__name__} ({value!r}) — iterating it would split "
                    "it into individual characters"
                )
        self._data: list[list[object]] = [list(row) for row in data]
        self._column_headers: list[str] = list(column_headers)
        self._row_labels: list[str] = list(row_labels)
        self._corner: str = corner
        self._name: str = name
        self._style: TableStyle = style if style is not None else TableStyle.DEFAULT
        self._column_types: list[ColumnType] = [ColumnType.ANY] * len(self._column_headers)
        if column_types is not None:
            if not isinstance(column_types, Mapping):
                raise TypeError(
                    f"column_types must be a mapping of column header to ColumnType, "
                    f"got {type(column_types).__name__}"
                )
            for header, column_type in column_types.items():
                if not isinstance(column_type, ColumnType):
                    raise TypeError(
                        f"column_types[{header!r}] must be a ColumnType, "
                        f"got {type(column_type).__name__}"
                    )
                try:
                    index = self._column_headers.index(header)
                except ValueError:
                    raise KeyError(f"no column headed {header!r}") from None
                self._column_types[index] = column_type
        # Set only by create()/read()/write() — a freshly built Table has no
        # sheet or history until it's actually placed somewhere.
        self._sheet: str | None = None
        self._created_at: datetime | None = None
        self._modified_at: datetime | None = None
        # What this object last saw on disk, for merging concurrent writers — see write().
        # Also which base row/column each current one descends from (None: added since),
        # since a rename is the one change a diff can't tell from a delete plus an add.
        self._base: Snapshot | None = None
        self._row_origin: dict[str, str | None] = {}
        self._column_origin: dict[str, str | None] = {}
        self._last_merge: MergeReport | None = None
        self._validate()

    def _validate(self) -> None:
        if not isinstance(self._name, str):
            raise TypeError(f"name must be str, got {type(self._name).__name__}: {self._name!r}")
        if not self._name:
            raise ValueError("name must not be empty")
        for label in self._row_labels:
            if not isinstance(label, str):
                raise TypeError(f"row labels must be str, got {type(label).__name__}: {label!r}")
            if not label:
                raise ValueError("row labels must not be empty")
        for header in self._column_headers:
            if not isinstance(header, str):
                raise TypeError(
                    f"column headers must be str, got {type(header).__name__}: {header!r}"
                )
            if not header:
                raise ValueError("column headers must not be empty")
        if not self._column_headers:
            raise ValueError("column_headers must not be empty")
        if not isinstance(self._style, TableStyle):
            raise TypeError(f"style must be a TableStyle, got {type(self._style).__name__}")
        if not isinstance(self._corner, str):
            got = type(self._corner).__name__
            raise TypeError(f"corner must be str, got {got}: {self._corner!r}")
        if self._row_labels and len(self._row_labels) != len(self._data):
            raise ValueError(f"{len(self._data)} data rows but {len(self._row_labels)} row labels")
        if self._data and not self._row_labels:
            raise ValueError(
                "this table has data but no row labels — give row_labels, or use the "
                "grid layout for positional-only data"
            )
        width = len(self._column_headers)
        for i, row in enumerate(self._data):
            if len(row) != width:
                raise ValueError(
                    f"data row {i} has {len(row)} values but there are {width} column headers"
                )

    # ------------------------------------------------------------------ read

    @classmethod
    def read(cls, path: str | Path, name: str) -> Table:
        """Read the named table called *name*.

        If the reserved schema sheet is missing entirely, it's automatically
        rebuilt by scanning the workbook for table markers first — see
        :class:`~pyhandlexl.errors.SchemaRebuiltWarning`. This can make a
        `read` write to the file, the same way a self-heal can.

        The returned table remembers what it read, so a later :meth:`write`
        can merge with anything another writer saved in between.

        Raises:
            TableNotFoundError: no such table exists, or its marker cannot be
                found on its recorded sheet.
        """
        workbook = safe_load(path)
        try:
            schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
            entries = mt.load_schema(workbook)
            entry = mt.get_entry(entries, name)
            located = mt.verify_or_locate(workbook, entry)
            healed = located != entry
            snapshot = mt.read_snapshot(workbook[located.sheet], located)

            if healed:
                entries[name] = located
            if healed or not schema_existed:
                mt.save_schema(workbook, entries)
                atomic_save(workbook, path)
        finally:
            workbook.close()

        table = cls(
            snapshot.data,
            snapshot.row_labels,
            snapshot.corner,
            column_headers=snapshot.column_headers,
            name=name,
            style=located.style,
            column_types=dict(zip(snapshot.column_headers, located.column_types, strict=True)),
        )
        table._sheet = located.sheet
        table._created_at = located.created_at
        table._modified_at = located.modified_at
        table._set_base()

        return table

    @classmethod
    def from_dict(
        cls,
        table: Mapping[str, Mapping[str, object]],
        corner: str = "",
        *,
        name: str,
        column_types: Mapping[str, ColumnType] | None = None,
    ) -> Table:
        """Build a Table from a dict of dicts: row label -> {column header: value}.

        The outer keys become the row labels, in order. Every inner dict must
        have the same keys in the same order — that order becomes the column
        headers (``ValueError`` otherwise). *table* must not be empty — with
        no rows there's nothing to infer the column headers from; construct
        the table directly and pass ``column_headers=`` for that.
        """
        if not isinstance(table, Mapping):
            raise TypeError(
                f"table must be a mapping of row label to {{column header: value}}, "
                f"got {type(table).__name__}"
            )
        row_labels = list(table.keys())
        rows = []
        for label, row in table.items():
            if not isinstance(row, Mapping):
                raise TypeError(
                    f"row {label!r} must be a mapping of column header to value, "
                    f"got {type(row).__name__}"
                )
            rows.append(dict(row))
        headers = list(rows[0].keys()) if rows else []
        for label, row in zip(row_labels, rows, strict=True):
            if list(row.keys()) != headers:
                raise ValueError(
                    f"row {label!r} has keys {list(row.keys())!r}, expected {headers!r}"
                )
        data = [list(row.values()) for row in rows]
        return cls(
            data, row_labels, corner, column_headers=headers, name=name, column_types=column_types
        )

    # ------------------------------------------------------------ properties

    @property
    def name(self) -> str:
        """This table's name."""
        return self._name

    def set_corner(self, value: str) -> None:
        """Change the value of cell A1. Read it back through ``t.data.corner``."""
        if not isinstance(value, str):
            raise TypeError(f"corner must be str, got {type(value).__name__}: {value!r}")
        self._corner = value

    @property
    def style(self) -> TableStyle:
        """This table's visual style — defaults to ``TableStyle.DEFAULT``.

        Settable. A change takes effect (and is persisted) on the next
        :meth:`create`/:meth:`write` — it repaints the table's full current
        region, so it has no visible effect until then.
        """
        return self._style

    @style.setter
    def style(self, value: TableStyle) -> None:
        if not isinstance(value, TableStyle):
            raise TypeError(f"style must be a TableStyle, got {type(value).__name__}")
        self._style = value

    @property
    def column_types(self) -> dict[str, ColumnType]:
        """Each column header's type restriction — ``ColumnType.ANY`` (no
        restriction) unless declared otherwise.

        Change one column's restriction with :meth:`set_column_type`.
        """
        return dict(zip(self._column_headers, self._column_types, strict=True))

    def set_column_type(self, header: str, column_type: ColumnType) -> None:
        """Restrict *header*'s column to *column_type*.

        ``ColumnType.ANY`` removes any existing restriction. Checked (along
        with every other column's) on the next :meth:`create`/:meth:`write`
        — ``ColumnTypeError`` if a data value already in the column doesn't
        match; existing data isn't checked until then.
        """
        if not isinstance(column_type, ColumnType):
            raise TypeError(f"column_type must be a ColumnType, got {type(column_type).__name__}")
        self._column_types[self._column_index(header)] = column_type

    @property
    def last_merge(self) -> MergeReport | None:
        """What the most recent :meth:`write` merged in from other writers, or ``None``.

        ``None`` if that write found nothing had changed on disk since this table
        was read (or if it hasn't written yet). Otherwise a
        :class:`~pyhandlexl._merge.MergeReport`: the rows and columns other writers
        added or removed, how many of your cells took their value, and any
        ``conflicts`` where yours won (also raised as a ``MergeConflictWarning``).
        """
        return self._last_merge

    def _snapshot(self) -> Snapshot:
        return Snapshot(
            list(self._row_labels),
            list(self._column_headers),
            [list(row) for row in self._data],
            self._corner,
            self._style,
            list(self._column_types),
        )

    def _adopt(self, snapshot: Snapshot) -> None:
        self._row_labels = list(snapshot.row_labels)
        self._column_headers = list(snapshot.column_headers)
        self._data = [list(row) for row in snapshot.data]
        self._corner = snapshot.corner
        self._style = snapshot.style
        self._column_types = list(snapshot.column_types)

    def _set_base(self) -> None:
        """Remember the current state as what's on disk, for the next write's merge."""
        self._base = self._snapshot()
        self._row_origin = {label: label for label in self._row_labels}
        self._column_origin = {header: header for header in self._column_headers}

    @property
    def info(self) -> TableInfo:
        """A read-only snapshot of this table's metadata — see :class:`TableInfo`."""
        return TableInfo(
            created_at=self._created_at,
            modified_at=self._modified_at,
            n_rows=len(self._data),
            n_cols=self._width(),
            sheet=self._sheet,
            style=self._style,
            column_types=self.column_types,
        )

    @property
    def data(self) -> TableData:
        """A snapshot of the table's rows, columns, row labels, column headers, and corner."""
        width = max((len(row) for row in self._data), default=0)
        padded = [row + [None] * (width - len(row)) for row in self._data]
        return TableData(
            rows=[list(row) for row in self._data],
            columns=[list(column) for column in zip(*padded, strict=True)],
            row_labels=list(self._row_labels),
            column_headers=list(self._column_headers),
            corner=self._corner,
        )

    def to_dict(self) -> dict[str, dict[str, object]]:
        """The table as a dict of dicts: row label -> {column header: value}.

        Preserves both axes in one structure — the inverse of
        :meth:`from_dict`. A duplicate row label or column header (only
        possible on a table read from a file) collapses to its last value,
        since dict keys must be unique.
        """
        return {
            label: dict(zip(self._column_headers, row, strict=True))
            for label, row in zip(self._row_labels, self._data, strict=True)
        }

    # ---------------------------------------------------------- label access

    def _row_index(self, label: str) -> int:
        try:
            return self._row_labels.index(label)
        except ValueError:
            raise KeyError(f"no row labeled {label!r}") from None

    def _column_index(self, header: str) -> int:
        try:
            return self._column_headers.index(header)
        except ValueError:
            raise KeyError(f"no column headed {header!r}") from None

    def read_row(self, row_label: str) -> list[object]:
        """The data row for *row_label* (labels not included)."""
        return list(self._data[self._row_index(row_label)])

    def read_column(self, column_header: str) -> list[object]:
        """The data column for *column_header* (headers not included)."""
        j = self._column_index(column_header)
        return [row[j] for row in self._data]

    # ------------------------------------------------- position/label access

    def _width(self) -> int:
        return len(self._column_headers)

    def _dimensions(self) -> tuple[int, int, int, int]:
        # Column headers (and so the header row) always exist; the label
        # column only once the table actually has a labeled row.
        col_start = 1 if self._row_labels else 0
        n_rows = 1 + len(self._data)
        n_cols = col_start + self._width()
        return 1, col_start, n_rows, n_cols

    def _classify(self, row: int, col_num: int) -> tuple[str, int, int]:
        """Return ``(kind, i, j)`` where *kind* is corner/header/label/data.

        *i* and *j* are indices into the relevant list (``_data`` is indexed
        by both).
        """
        row_start, col_start, n_rows, n_cols = self._dimensions()
        if not (1 <= row <= n_rows and 1 <= col_num <= n_cols):
            raise IndexError(f"cell ({row}, {col_num}) is outside the {n_rows}x{n_cols} table")

        in_header_row = row == 1
        in_label_column = bool(self._row_labels) and col_num == 1
        if in_header_row and in_label_column:
            return "corner", 0, 0
        if in_header_row:
            return "header", 0, col_num - 1 - col_start
        if in_label_column:
            return "label", row - 1 - row_start, 0
        return "data", row - 1 - row_start, col_num - 1 - col_start

    def _dispatch(
        self, ref: str | None, row: int | str | None, column: int | str | None
    ) -> tuple[bool, int | str, int | str]:
        """Resolve ``ref``/``row``/``column`` into ``(is_position, row_val, col_val)``.

        ``is_position=True`` — ``row_val``/``col_val`` are 1-based Excel ints.
        ``is_position=False`` — they are a row label / column header string.
        """
        if ref is not None:
            if row is not None or column is not None:
                raise TypeError("give either ref or row=/column=, not both")
            try:
                r, c = coordinate_to_tuple(ref)
            except (ValueError, TypeError, UnboundLocalError):
                raise ValueError(f"invalid cell reference: {ref!r}") from None
            return True, r, c

        if row is None or column is None:
            raise TypeError("give a ref like 'B2', or both row= and column=")

        if isinstance(row, int) and isinstance(column, int):
            return True, row, column
        if isinstance(row, str) and isinstance(column, str):
            return False, row, column
        raise TypeError("row and column must both be int (position) or both be str (label)")

    def read_cell(
        self,
        ref: str | None = None,
        *,
        row: int | str | None = None,
        column: int | str | None = None,
    ) -> object:
        """A single value, addressed either by position or by label.

        Give either a ref like ``"B2"``, or ``row=``/``column=`` as a matching
        pair: both ints for a 1-based position within the table (row 1 is the
        header row, column 1 is the label column, so ``read_cell(row=2,
        column=2)`` is the first data cell), or both strings for a row label /
        column header pair — e.g. ``read_cell(row="Alice", column="q1")``.

        By position, ``read_cell`` can reach any cell — header, label, corner,
        or data. By label it always reads data (the label-addressed equivalent
        of ``read_cell(row=2, column=2)``, wherever that intersection lives).
        """
        is_position, r, c = self._dispatch(ref, row, column)
        if is_position:
            kind, i, j = self._classify(r, c)
            if kind == "corner":
                return self._corner
            if kind == "header":
                return self._column_headers[j]
            if kind == "label":
                return self._row_labels[i]
            return self._data[i][j]
        return self._data[self._row_index(r)][self._column_index(c)]

    # -------------------------------------------------------------- mutation

    def set_cell(
        self,
        ref: str | None = None,
        *,
        row: int | str | None = None,
        column: int | str | None = None,
        value: object,
    ) -> None:
        """Set a single **data** value, addressed by position or by label.

        Same addressing as :meth:`read_cell`. Only ever sets data — by
        position, addressing a header, row label, or the corner raises
        ``ValueError``; use :meth:`rename_column`, :meth:`rename_row`, or
        :meth:`set_corner` for those. By label there's no other kind of cell
        to reach, so it always sets data.
        """
        is_position, r, c = self._dispatch(ref, row, column)
        if is_position:
            self._set_by_position(r, c, value)
        else:
            self._data[self._row_index(r)][self._column_index(c)] = value

    def _set_by_position(self, row: int, col_num: int, value: object) -> None:
        kind, i, j = self._classify(row, col_num)
        if kind == "corner":
            raise ValueError("that cell is the corner — set it with table.set_corner(value)")
        if kind == "header":
            raise ValueError("that cell is a column header — rename it with rename_column()")
        if kind == "label":
            raise ValueError("that cell is a row label — rename it with rename_row()")
        self._data[i][j] = value

    def set_row(self, row_label: str, values: Iterable[object]) -> None:
        """Replace the data row for *row_label*; ``len(values)`` must match the column count."""
        new_row = list(values)
        i = self._row_index(row_label)
        expected = self._width()
        if len(new_row) != expected:
            raise ValueError(f"expected {expected} values, got {len(new_row)}")
        self._data[i] = new_row

    def set_column(self, column_header: str, values: Iterable[object]) -> None:
        """Replace the data column for *column_header*; ``len(values)`` must match the row count."""
        new_col = list(values)
        j = self._column_index(column_header)
        if len(new_col) != len(self._data):
            raise ValueError(f"expected {len(self._data)} values, got {len(new_col)}")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row[j] = value

    def _check_new_row(self, label: str, values: list[object]) -> None:
        if not isinstance(label, str):
            raise TypeError(f"row label must be str, got {type(label).__name__}: {label!r}")
        if not label:
            raise ValueError("row label must not be empty")
        if label in self._row_labels:
            raise ValueError(f"row label {label!r} already exists")
        if len(values) != self._width():
            raise ValueError(f"expected {self._width()} values, got {len(values)}")

    def add_row(self, label: str, values: Iterable[object]) -> None:
        """Append a labeled data row.

        ``len(values)`` must match the column count, and *label* must not
        already be in use (``ValueError``).
        """
        self.insert_row(len(self._data) + 1, label, values)

    def insert_row(self, position: int, label: str, values: Iterable[object]) -> None:
        """Insert a labeled data row before data-row *position* (1-based).

        ``position`` ranges from 1 (new first row) through
        ``len(data.rows) + 1`` — that top end inserts it as the last row, the
        same result as :meth:`add_row`. Same constraints as :meth:`add_row`:
        *label* must not already be in use, and ``values`` must match the
        column count.
        """
        if not isinstance(position, int):
            raise TypeError(f"position must be int, got {type(position).__name__}: {position!r}")
        new_row = list(values)
        self._check_new_row(label, new_row)
        n = len(self._data)
        if not 1 <= position <= n + 1:
            raise IndexError(f"position {position} is out of range for {n} rows (1..{n + 1})")
        self._data.insert(position - 1, new_row)
        self._row_labels.insert(position - 1, label)
        self._row_origin[label] = None

    def _check_new_column(self, header: str, values: list[object]) -> None:
        if not isinstance(header, str):
            raise TypeError(f"column header must be str, got {type(header).__name__}: {header!r}")
        if not header:
            raise ValueError("column header must not be empty")
        if header in self._column_headers:
            raise ValueError(f"column header {header!r} already exists")
        if len(values) != len(self._data):
            raise ValueError(f"expected {len(self._data)} values, got {len(values)}")

    def add_column(self, header: str, values: Iterable[object]) -> None:
        """Append a labeled data column.

        ``len(values)`` must match the row count, and *header* must not
        already be in use (``ValueError``).
        """
        self.insert_column(len(self._column_headers) + 1, header, values)

    def insert_column(self, position: int, header: str, values: Iterable[object]) -> None:
        """Insert a labeled data column before column-position *position* (1-based).

        ``position`` ranges from 1 (new first column) through
        ``len(data.column_headers) + 1`` — that top end inserts it as the
        last column, the same result as :meth:`add_column`. Same constraints
        as :meth:`add_column`: *header* must not already be in use, and
        ``values`` must match the row count.
        """
        if not isinstance(position, int):
            raise TypeError(f"position must be int, got {type(position).__name__}: {position!r}")
        new_col = list(values)
        self._check_new_column(header, new_col)
        n = len(self._column_headers)
        if not 1 <= position <= n + 1:
            raise IndexError(f"position {position} is out of range for {n} columns (1..{n + 1})")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row.insert(position - 1, value)
        self._column_headers.insert(position - 1, header)
        self._column_types.insert(position - 1, ColumnType.ANY)
        self._column_origin[header] = None

    def drop_row(self, label: str) -> None:
        """Remove the row labeled *label*."""
        i = self._row_index(label)
        del self._data[i]
        del self._row_labels[i]
        self._row_origin.pop(label, None)

    def drop_column(self, header: str) -> None:
        """Remove the column headed *header*.

        Raises ``ValueError`` if it's the only column left — a table always
        has at least one column header.
        """
        j = self._column_index(header)
        if len(self._column_headers) == 1:
            raise ValueError("cannot drop the only remaining column")
        del self._column_headers[j]
        del self._column_types[j]
        self._column_origin.pop(header, None)
        for data_row in self._data:
            del data_row[j]

    def rename_row(self, old: str, new: str) -> None:
        """Change a row label; *new* must not already be in use (``ValueError``)."""
        if not isinstance(new, str):
            raise TypeError(f"row label must be str, got {type(new).__name__}: {new!r}")
        if not new:
            raise ValueError("row label must not be empty")
        i = self._row_index(old)
        if new != old and new in self._row_labels:
            raise ValueError(f"row label {new!r} already exists")
        self._row_labels[i] = new
        if new != old:
            self._row_origin[new] = self._row_origin.pop(old, None)

    def rename_column(self, old: str, new: str) -> None:
        """Change a column header; *new* must not already be in use (``ValueError``)."""
        if not isinstance(new, str):
            raise TypeError(f"column header must be str, got {type(new).__name__}: {new!r}")
        if not new:
            raise ValueError("column header must not be empty")
        j = self._column_index(old)
        if new != old and new in self._column_headers:
            raise ValueError(f"column header {new!r} already exists")
        self._column_headers[j] = new
        if new != old:
            self._column_origin[new] = self._column_origin.pop(old, None)

    # --------------------------------------------------------------- display

    def show(self, *, rows: int | None = None, head: int | None = 5, tail: int | None = 5) -> None:
        """Print the table to the console as a plain aligned grid.

        By default, prints the first 5 and last 5 rows with a ``...`` divider
        between them (nothing is hidden if the table has 10 rows or fewer).
        ``rows=n`` overrides that and prints only the first *n* data rows.
        To print every row, pass ``head=None, tail=None`` explicitly. This is
        a console convenience only — it has nothing to do with cell formatting
        in the workbook.
        """
        for param_name, value in (("rows", rows), ("head", head), ("tail", tail)):
            if value is None:
                continue
            if not isinstance(value, int):
                raise TypeError(f"{param_name} must be int, got {type(value).__name__}: {value!r}")
            if value < 0:
                raise ValueError(f"{param_name} must not be negative, got {value}")

        data = self._data
        labels = self._row_labels
        divider_after: int | None = None

        if rows is not None:
            data = data[:rows]
            labels = labels[:rows] if labels else labels
        elif head is not None or tail is not None:
            h, t = head or 0, tail or 0
            if h + t >= len(data):
                pass
            else:
                data = data[:h] + data[len(data) - t :]
                labels = labels[:h] + labels[len(labels) - t :] if labels else labels
                divider_after = h

        grid: list[list[str]] = [[self._corner, *self._column_headers]]
        for i, row in enumerate(data):
            label = labels[i] if i < len(labels) else ""
            grid.append([label, *("" if v is None else str(v) for v in row)])
            if divider_after is not None and i == divider_after - 1 and divider_after < len(data):
                grid.append(["..." for _ in grid[-1]])

        widths = [max(len(row[c]) for row in grid) for c in range(len(grid[0]))]
        for row in grid:
            print("  ".join(cell.ljust(w) for cell, w in zip(row, widths, strict=True)))

    # ----------------------------------------------------------------- write

    def _assemble(self) -> list[list[object]]:
        # column_headers is always non-empty and, whenever there's data,
        # row_labels always matches it in length — both enforced by
        # _validate() and preserved by every mutator — so the corner and
        # every row label are always there to use.
        grid: list[list[object]] = [[self._corner, *self._column_headers]]
        for i, data_row in enumerate(self._data):
            grid.append([self._row_labels[i], *data_row])
        return grid

    def _check_column_types(self) -> None:
        for j, column_type in enumerate(self._column_types):
            if column_type is ColumnType.ANY:
                continue
            header = self._column_headers[j]
            for i, row in enumerate(self._data):
                value = row[j]
                if not column_type.allows(value):
                    raise ColumnTypeError(
                        f"column {header!r} is restricted to {column_type.value} — row "
                        f"{self._row_labels[i]!r} has {type(value).__name__} {value!r}"
                    )

    def write(self, path: str | Path) -> None:
        """Write this table back to its tracked position.

        The table must already exist — create it first with :meth:`create`
        (``TableNotFoundError`` otherwise). Growing it in place — more columns
        than it had before — shifts every table to its right on the same
        sheet to make room; growing rows never shifts anything.

        Every data value must be a type Excel can store (``CellTypeError``
        otherwise), and must match any column type restriction
        (``ColumnTypeError`` otherwise) — both checks happen here, not when
        values are set.

        **Concurrent writers.** If another writer changed this table on disk
        since this object last read or wrote it, their changes are *merged*
        with yours rather than overwritten: rows and columns they added appear
        in your table, edits to different cells both survive, and both sides
        appending just works. Where you both changed the very same thing
        differently, yours wins and a ``MergeConflictWarning`` lists what it
        overwrote. Afterwards this object holds the merged table, and
        :attr:`last_merge` says what came in. A table that was never read or
        created by this object has nothing to merge against, so it simply
        overwrites.

        The merge covers the time between your ``read`` and your ``write``. It
        is not a lock: two writes landing at the very same instant can still
        overwrite one another, since the check for changes and the save are
        separate steps.

        Raises:
            TableNotFoundError: the table doesn't exist (or was deleted).
            CellTypeError / ColumnTypeError: a value isn't allowed — checked
                *after* merging, so a value someone else's new restriction
                rules out is caught here too; nothing is written and this
                object is left exactly as it was.
            FileLockedError: the file stayed locked (open in Excel) through every retry.
        """
        report: MergeReport | None = None
        workbook = safe_load(path)
        try:
            entries = mt.load_schema(workbook)
            entry = mt.get_entry(entries, self._name)
            entry = mt.verify_or_locate(workbook, entry)
            entries[self._name] = entry
            self._sheet = entry.sheet
            self._created_at = entry.created_at

            before = self._snapshot()
            if self._base is not None:
                theirs = mt.read_snapshot(workbook[entry.sheet], entry)
                if not same_state(theirs, self._base):
                    merged, report = merge(
                        self._base, before, self._row_origin, self._column_origin, theirs
                    )
                    self._adopt(merged)

            try:
                assembled = self._assemble()
                for row in assembled:
                    for value in row:
                        check_cell_value(value)
                self._check_column_types()

                now = datetime.now(timezone.utc)
                new_n_rows = len(self._data)
                new_n_cols = self._width()

                growth = new_n_cols - entry.n_cols
                if growth > 0:
                    neighbours = mt.tables_to_shift(entries, entry)
                    if neighbours:
                        rightmost = max(e.anchor_col + e.width - 1 for e in neighbours) + growth
                        check_dimensions(1, rightmost)
                    mt.shift_right(workbook, entries, entry, growth)

                check_dimensions(entry.anchor_row + new_n_rows + 1, entry.anchor_col + new_n_cols)

                ws = workbook[entry.sheet]
                old_height, old_width = entry.height, entry.width
                new_height, new_width = new_n_rows + 2, new_n_cols + 1
                mt.clear_region(
                    ws,
                    entry.anchor_row,
                    entry.anchor_col,
                    max(old_height, new_height),
                    max(old_width, new_width),
                )
                mt.write_region(ws, entry.anchor_row, entry.anchor_col, [[mt.MARKER, self._name]])
                mt.write_region(ws, entry.anchor_row + 1, entry.anchor_col, assembled)

                updated_entry = replace(
                    entry,
                    n_rows=new_n_rows,
                    n_cols=new_n_cols,
                    style=self._style,
                    column_types=list(self._column_types),
                    modified_at=now,
                )
                entries[self._name] = updated_entry
                mt.paint_table(ws, updated_entry)
                mt.save_schema(workbook, entries)
                atomic_save(workbook, path)
            except BaseException:
                self._adopt(before)  # leave this object exactly as the caller had it
                raise
        finally:
            workbook.close()

        self._modified_at = now
        self._set_base()
        self._last_merge = report
        if report is not None and report.conflicts:
            warnings.warn(MergeConflictWarning(_conflict_message(self._name, report)), stacklevel=2)

    def create(self, path: str | Path, sheet: str) -> None:
        """Place this table as a brand-new table on *sheet*.

        Tables on a sheet stack left to right with one empty column between
        them, always starting at row 1.

        Raises:
            TableExistsError: a table named this already exists in the workbook.
            SheetNotFoundError: *sheet* does not exist.
            SheetKindError: *sheet* already holds plain grid data (see
                ``write_sheet``/``append_rows``), or is the reserved schema
                sheet.
            CellTypeError: a data value is not a type Excel can store.
            ColumnTypeError: a data value doesn't match its column's type restriction.
            DimensionError: the placed table would exceed Excel's grid limits.
            FileLockedError: the file stayed locked (open in Excel) through every retry.
        """
        workbook = safe_load(path)
        try:
            if sheet == mt.SCHEMA_SHEET:
                raise SheetKindError(f"{mt.SCHEMA_SHEET!r} is reserved and cannot hold a table")
            if sheet not in workbook.sheetnames:
                raise SheetNotFoundError(sheet)
            entries = mt.load_schema(workbook)
            if mt.sheet_kind(workbook, entries, sheet) == "grid":
                raise SheetKindError(
                    f"sheet {sheet!r} holds plain grid data — placing a table there would "
                    "corrupt it; use a different sheet, or clear_all_sheet_data() first"
                )
            mt.check_not_exists(entries, self._name)

            assembled = self._assemble()
            for row in assembled:
                for value in row:
                    check_cell_value(value)
            self._check_column_types()

            anchor_row, anchor_col = mt.find_placement(entries, sheet)
            n_rows = len(self._data)
            n_cols = self._width()
            check_dimensions(anchor_row + n_rows + 1, anchor_col + n_cols)

            ws = workbook[sheet]
            mt.write_region(ws, anchor_row, anchor_col, [[mt.MARKER, self._name]])
            mt.write_region(ws, anchor_row + 1, anchor_col, assembled)

            now = datetime.now(timezone.utc)
            entry = mt.TableEntry(
                self._name,
                sheet,
                anchor_row,
                anchor_col,
                n_rows,
                n_cols,
                self._style,
                list(self._column_types),
                created_at=now,
                modified_at=now,
            )
            entries[self._name] = entry
            mt.paint_table(ws, entry)
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)
        finally:
            workbook.close()

        self._sheet = sheet
        self._created_at = now
        self._modified_at = now
        self._set_base()
        self._last_merge = None

    # --------------------------------------------------------------- dunders

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Table):
            return NotImplemented
        return (
            self._data == other._data
            and self._column_headers == other._column_headers
            and self._row_labels == other._row_labels
            and self._corner == other._corner
            and self._column_types == other._column_types
        )

    def __repr__(self) -> str:
        return (
            f"Table(name={self._name!r}, rows={len(self._data)}, "
            f"columns={len(self._column_headers)}, corner={self._corner!r})"
        )
