"""The Table class: a worksheet read as column headers, row labels, and a data grid.

Layout convention: row 1 holds the column headers, column A holds the row
labels, cell A1 is the "corner", and the data region is everything from B2
onward. Row labels and column headers are always strings. Positional access
uses Excel coordinates (row 1 is the header row, column 1 is the label column).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import coordinate_to_tuple

from pyhandlexl.core import read_sheet, write_sheet


@dataclass(frozen=True)
class TableData:
    """A read-only snapshot of a Table's content."""

    rows: list[list[object]]
    columns: list[list[object]]
    row_labels: list[str]
    column_headers: list[str]
    corner: object


class Table:
    """A labeled table backed by a worksheet.

    Construct directly from parts, or with :meth:`read` from a file. Mutation
    methods edit the table in place and return ``None``. Row labels and column
    headers are always ``str``.
    """

    def __init__(
        self,
        data: Iterable[Iterable[object]],
        column_headers: Iterable[str] = (),
        row_labels: Iterable[str] = (),
        corner: object = "",
    ) -> None:
        self._data: list[list[object]] = [list(row) for row in data]
        self._column_headers: list[str] = list(column_headers)
        self._row_labels: list[str] = list(row_labels)
        self._corner: object = corner
        self._validate()

    def _validate(self) -> None:
        for label in self._row_labels:
            if not isinstance(label, str):
                raise TypeError(f"row labels must be str, got {type(label).__name__}: {label!r}")
        for header in self._column_headers:
            if not isinstance(header, str):
                raise TypeError(
                    f"column headers must be str, got {type(header).__name__}: {header!r}"
                )
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
    def read(
        cls,
        path: str | Path,
        sheet: str | None = None,
        *,
        column_headers: bool = True,
        row_labels: bool = True,
    ) -> Table:
        """Read a worksheet into a Table.

        Args:
            path: the .xlsx file.
            sheet: worksheet name, or ``None`` for the active sheet.
            column_headers: treat row 1 as column headers.
            row_labels: treat column A as row labels.
        """
        grid = read_sheet(path, sheet, pad=True)
        if not grid:
            return cls([], [], [], "")

        row_start = 1 if column_headers else 0
        col_start = 1 if row_labels else 0

        corner = grid[0][0] if (column_headers and row_labels) else ""
        headers = grid[0][col_start:] if column_headers else []
        labels = [grid[r][0] for r in range(row_start, len(grid))] if row_labels else []
        data = [grid[r][col_start:] for r in range(row_start, len(grid))]

        return cls(data, headers, labels, corner)

    # ------------------------------------------------------------ properties

    @property
    def corner(self) -> object:
        """The value of cell A1. Settable — this is how you change it."""
        return self._corner

    @corner.setter
    def corner(self, value: object) -> None:
        self._corner = value

    @property
    def data(self) -> TableData:
        """A snapshot of the table's rows, columns, row labels, column headers, and corner."""
        return TableData(
            rows=[list(row) for row in self._data],
            columns=[list(column) for column in zip(*self._data, strict=True)],
            row_labels=list(self._row_labels),
            column_headers=list(self._column_headers),
            corner=self._corner,
        )

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
        pair: both ints for a 1-based Excel position (row 1 is the header row,
        column 1 is the label column, so ``read_cell(row=2, column=2)`` is the
        first data cell), or both strings for a row label / column header pair
        — e.g. ``read_cell(row="Alice", column="q1")``.

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

    def add_row(self, label: str, values: Iterable[object]) -> None:
        """Append a labeled data row; ``len(values)`` must match the column count."""
        if not isinstance(label, str):
            raise TypeError(f"row label must be str, got {type(label).__name__}: {label!r}")
        new_row = list(values)
        if self._data and not self._row_labels:
            raise ValueError("this table has no row labels; use the raw layer to add rows")
        if (self._data or self._column_headers) and len(new_row) != self._width():
            raise ValueError(f"expected {self._width()} values, got {len(new_row)}")
        self._data.append(new_row)
        self._row_labels.append(label)

    def add_column(self, header: str, values: Iterable[object]) -> None:
        """Append a labeled data column; ``len(values)`` must match the row count."""
        if not isinstance(header, str):
            raise TypeError(f"column header must be str, got {type(header).__name__}: {header!r}")
        new_col = list(values)
        if any(self._data) and not self._column_headers:
            raise ValueError("this table has no column headers; use the raw layer to add columns")
        if len(new_col) != len(self._data):
            raise ValueError(f"expected {len(self._data)} values, got {len(new_col)}")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row.append(value)
        self._column_headers.append(header)

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
        """Change a row label."""
        if not isinstance(new, str):
            raise TypeError(f"row label must be str, got {type(new).__name__}: {new!r}")
        self._row_labels[self._row_index(old)] = new

    def rename_column(self, old: str, new: str) -> None:
        """Change a column header."""
        if not isinstance(new, str):
            raise TypeError(f"column header must be str, got {type(new).__name__}: {new!r}")
        self._column_headers[self._column_index(old)] = new

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

    def write(self, path: str | Path, sheet: str | None = None) -> None:
        """Write the table to *path*, reassembling headers into row 1 and labels into column A."""
        write_sheet(path, self._assemble(), sheet)

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
            f"Table(rows={len(self._data)}, columns={len(self._column_headers)}, "
            f"corner={self._corner!r})"
        )
