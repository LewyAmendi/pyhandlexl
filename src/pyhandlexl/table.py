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

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, replace
from pathlib import Path

from openpyxl.utils import coordinate_to_tuple

from pyhandlexl import _multi_table as mt
from pyhandlexl._safety import atomic_save, safe_load
from pyhandlexl.errors import SheetNotFoundError
from pyhandlexl.validate import check_cell_value, check_dimensions


def _to_label(value: object) -> str:
    """Coerce a header/label/corner cell to ``str``; an empty cell becomes ``""``."""
    return "" if value is None else str(value)


@dataclass(frozen=True)
class TableData:
    """A read-only snapshot of a Table's content."""

    rows: list[list[object]]
    columns: list[list[object]]
    row_labels: list[str]
    column_headers: list[str]
    corner: str


class Table:
    """A named table, always with column headers, row labels, and a corner.

    Construct directly from parts (``name=`` required), or with :meth:`read`
    from a file. Mutation methods edit the table in place and return
    ``None``. Row labels and column headers are always ``str``.
    """

    def __init__(
        self,
        data: Iterable[Iterable[object]],
        column_headers: Iterable[str] = (),
        row_labels: Iterable[str] = (),
        corner: str = "",
        *,
        name: str,
    ) -> None:
        self._data: list[list[object]] = [list(row) for row in data]
        self._column_headers: list[str] = list(column_headers)
        self._row_labels: list[str] = list(row_labels)
        self._corner: str = corner
        self._name: str = name
        self._validate()

    def _validate(self) -> None:
        if not isinstance(self._name, str):
            raise TypeError(f"name must be str, got {type(self._name).__name__}: {self._name!r}")
        if not self._name:
            raise ValueError("name must not be empty")
        for label in self._row_labels:
            if not isinstance(label, str):
                raise TypeError(f"row labels must be str, got {type(label).__name__}: {label!r}")
        for header in self._column_headers:
            if not isinstance(header, str):
                raise TypeError(
                    f"column headers must be str, got {type(header).__name__}: {header!r}"
                )
        if not isinstance(self._corner, str):
            got = type(self._corner).__name__
            raise TypeError(f"corner must be str, got {got}: {self._corner!r}")
        if self._row_labels and len(self._row_labels) != len(self._data):
            raise ValueError(f"{len(self._data)} data rows but {len(self._row_labels)} row labels")
        if self._column_headers:
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

        Raises:
            TableNotFoundError: no such table exists, or its marker cannot be
                found on its recorded sheet.
        """
        workbook = safe_load(path)
        try:
            entries = mt.load_schema(workbook)
            entry = mt.get_entry(entries, name)
            located = mt.verify_or_locate(workbook, entry)
            healed = located != entry

            ws = workbook[located.sheet]
            block = mt.read_region(
                ws,
                located.anchor_row + 1,
                located.anchor_col,
                located.n_rows + 1,
                located.n_cols + 1,
            )
            corner = _to_label(block[0][0])
            headers = [_to_label(h) for h in block[0][1:]]
            labels = [_to_label(row[0]) for row in block[1:]]
            data = [list(row[1:]) for row in block[1:]]
            table = cls(data, headers, labels, corner, name=name)

            if healed:
                entries[name] = located
                mt.save_schema(workbook, entries)
                atomic_save(workbook, path)
            return table
        finally:
            workbook.close()

    @classmethod
    def from_dict(
        cls,
        table: Mapping[str, Mapping[str, object]],
        corner: str = "",
        *,
        name: str,
    ) -> Table:
        """Build a Table from a dict of dicts: row label -> {column header: value}.

        The outer keys become the row labels, in order. Every inner dict must
        have the same keys in the same order — that order becomes the column
        headers (``ValueError`` otherwise).
        """
        row_labels = list(table.keys())
        rows = [dict(row) for row in table.values()]
        headers = list(rows[0].keys()) if rows else []
        for label, row in zip(row_labels, rows, strict=True):
            if list(row.keys()) != headers:
                raise ValueError(
                    f"row {label!r} has keys {list(row.keys())!r}, expected {headers!r}"
                )
        data = [list(row.values()) for row in rows]
        return cls(data, headers, row_labels, corner, name=name)

    # ------------------------------------------------------------ properties

    @property
    def name(self) -> str:
        """This table's name."""
        return self._name

    @property
    def corner(self) -> str:
        """The value of cell A1 (always ``str``). Settable — this is how you change it."""
        return self._corner

    @corner.setter
    def corner(self, value: str) -> None:
        if not isinstance(value, str):
            raise TypeError(f"corner must be str, got {type(value).__name__}: {value!r}")
        self._corner = value

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
        if self._column_headers:
            return len(self._column_headers)
        return len(self._data[0]) if self._data else 0

    def _dimensions(self) -> tuple[int, int, int, int]:
        row_start = 1 if self._column_headers else 0
        col_start = 1 if self._row_labels else 0
        n_rows = row_start + len(self._data)
        n_cols = col_start + self._width()
        return row_start, col_start, n_rows, n_cols

    def _classify(self, row: int, col_num: int) -> tuple[str, int, int]:
        """Return ``(kind, i, j)`` where *kind* is corner/header/label/data.

        *i* and *j* are indices into the relevant list (``_data`` is indexed
        by both).
        """
        row_start, col_start, n_rows, n_cols = self._dimensions()
        if not (1 <= row <= n_rows and 1 <= col_num <= n_cols):
            raise IndexError(f"cell ({row}, {col_num}) is outside the {n_rows}x{n_cols} table")

        in_header_row = bool(self._column_headers) and row == 1
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
            r, c = coordinate_to_tuple(ref)
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
        ``ValueError``; use :meth:`rename_column`, :meth:`rename_row`, or the
        ``corner`` property for those. By label there's no other kind of cell
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
            raise ValueError("that cell is the corner — set it with table.corner = value")
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
        if label in self._row_labels:
            raise ValueError(f"row label {label!r} already exists")
        if self._data and not self._row_labels:
            raise ValueError("this table has no row labels; use the grid layout to add rows")
        if (self._data or self._column_headers) and len(values) != self._width():
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
        new_row = list(values)
        self._check_new_row(label, new_row)
        n = len(self._data)
        if not 1 <= position <= n + 1:
            raise IndexError(f"position {position} is out of range for {n} rows (1..{n + 1})")
        self._data.insert(position - 1, new_row)
        self._row_labels.insert(position - 1, label)

    def _check_new_column(self, header: str, values: list[object]) -> None:
        if not isinstance(header, str):
            raise TypeError(f"column header must be str, got {type(header).__name__}: {header!r}")
        if header in self._column_headers:
            raise ValueError(f"column header {header!r} already exists")
        if any(self._data) and not self._column_headers:
            raise ValueError("this table has no column headers; use the grid layout to add columns")
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
        new_col = list(values)
        self._check_new_column(header, new_col)
        n = len(self._column_headers)
        if not 1 <= position <= n + 1:
            raise IndexError(f"position {position} is out of range for {n} columns (1..{n + 1})")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row.insert(position - 1, value)
        self._column_headers.insert(position - 1, header)

    def drop_row(self, label: str) -> None:
        """Remove the row labeled *label*."""
        i = self._row_index(label)
        del self._data[i]
        del self._row_labels[i]

    def drop_column(self, header: str) -> None:
        """Remove the column headed *header*."""
        j = self._column_index(header)
        del self._column_headers[j]
        for data_row in self._data:
            del data_row[j]

    def rename_row(self, old: str, new: str) -> None:
        """Change a row label; *new* must not already be in use (``ValueError``)."""
        if not isinstance(new, str):
            raise TypeError(f"row label must be str, got {type(new).__name__}: {new!r}")
        i = self._row_index(old)
        if new != old and new in self._row_labels:
            raise ValueError(f"row label {new!r} already exists")
        self._row_labels[i] = new

    def rename_column(self, old: str, new: str) -> None:
        """Change a column header; *new* must not already be in use (``ValueError``)."""
        if not isinstance(new, str):
            raise TypeError(f"column header must be str, got {type(new).__name__}: {new!r}")
        j = self._column_index(old)
        if new != old and new in self._column_headers:
            raise ValueError(f"column header {new!r} already exists")
        self._column_headers[j] = new

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

        grid: list[list[str]] = []
        if self._column_headers:
            grid.append([self._corner, *self._column_headers])
        for i, row in enumerate(data):
            label = labels[i] if i < len(labels) else ""
            grid.append([label, *("" if v is None else str(v) for v in row)])
            if divider_after is not None and i == divider_after - 1 and divider_after < len(data):
                grid.append(["..." for _ in grid[-1]])

        if not grid:
            print("(empty table)")
            return
        widths = [max(len(row[c]) for row in grid) for c in range(len(grid[0]))]
        for row in grid:
            print("  ".join(cell.ljust(w) for cell, w in zip(row, widths, strict=True)))

    # ----------------------------------------------------------------- write

    def _assemble(self) -> list[list[object]]:
        grid: list[list[object]] = []
        if self._column_headers:
            header_row: list[object] = [self._corner] if self._row_labels else []
            header_row.extend(self._column_headers)
            grid.append(header_row)
        for i, data_row in enumerate(self._data):
            out_row: list[object] = [self._row_labels[i]] if self._row_labels else []
            out_row.extend(data_row)
            grid.append(out_row)
        return grid

    def write(self, path: str | Path) -> None:
        """Write this table back to its tracked position.

        The table must already exist — create it first with :meth:`create`
        (``TableNotFoundError`` otherwise). Growing it in place — more columns
        than it had before — shifts every table to its right on the same
        sheet to make room; growing rows never shifts anything.

        Every data value must be a type Excel can store (``CellTypeError``
        otherwise) — the check happens here, not when values are set.
        """
        workbook = safe_load(path)
        try:
            entries = mt.load_schema(workbook)
            entry = mt.get_entry(entries, self._name)
            entry = mt.verify_or_locate(workbook, entry)
            entries[self._name] = entry

            assembled = self._assemble()
            for row in assembled:
                for value in row:
                    check_cell_value(value)

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

            entries[self._name] = replace(entry, n_rows=new_n_rows, n_cols=new_n_cols)
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)
        finally:
            workbook.close()

    def create(self, path: str | Path, sheet: str) -> None:
        """Place this table as a brand-new table on *sheet*.

        Tables on a sheet stack left to right with one empty column between
        them, always starting at row 1.

        Raises:
            TableExistsError: a table named this already exists in the workbook.
            SheetNotFoundError: *sheet* does not exist.
            CellTypeError: a data value is not a type Excel can store.
            DimensionError: the placed table would exceed Excel's grid limits.
        """
        workbook = safe_load(path)
        try:
            if sheet not in workbook.sheetnames:
                raise SheetNotFoundError(sheet)
            entries = mt.load_schema(workbook)
            mt.check_not_exists(entries, self._name)

            assembled = self._assemble()
            for row in assembled:
                for value in row:
                    check_cell_value(value)

            anchor_row, anchor_col = mt.find_placement(entries, sheet)
            n_rows = len(self._data)
            n_cols = self._width()
            check_dimensions(anchor_row + n_rows + 1, anchor_col + n_cols)

            ws = workbook[sheet]
            mt.write_region(ws, anchor_row, anchor_col, [[mt.MARKER, self._name]])
            mt.write_region(ws, anchor_row + 1, anchor_col, assembled)

            entries[self._name] = mt.TableEntry(
                self._name, sheet, anchor_row, anchor_col, n_rows, n_cols
            )
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)
        finally:
            workbook.close()

    # --------------------------------------------------------------- dunders

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Table):
            return NotImplemented
        return (
            self._data == other._data
            and self._column_headers == other._column_headers
            and self._row_labels == other._row_labels
            and self._corner == other._corner
        )

    def __repr__(self) -> str:
        return (
            f"Table(name={self._name!r}, rows={len(self._data)}, "
            f"columns={len(self._column_headers)}, corner={self._corner!r})"
        )
