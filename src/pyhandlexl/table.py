"""The Table class: a worksheet read as column headers, row labels, and a data grid.

Layout convention: row 1 holds the column headers, column A holds the row
labels, cell A1 is the "corner", and the data region is everything from B2
onward. Positional access uses Excel coordinates (row 1 is the header row,
column 1 is the label column).
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from pathlib import Path

from openpyxl.utils import column_index_from_string, coordinate_to_tuple

from pyhandlexl.core import read_sheet, write_sheet


class Table:
    """A labeled table backed by a worksheet.

    Construct directly from parts, or with :meth:`read` from a file. Mutation
    methods edit the table in place and return ``None``.
    """

    def __init__(
        self,
        data: Iterable[Iterable[object]],
        column_headers: Iterable[object] = (),
        row_labels: Iterable[object] = (),
        corner: object = "",
    ) -> None:
        self._data: list[list[object]] = [list(row) for row in data]
        self._column_headers: list[object] = list(column_headers)
        self._row_labels: list[object] = list(row_labels)
        self._corner: object = corner
        self._validate()

    def _validate(self) -> None:
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
        """The value of cell A1."""
        return self._corner

    @corner.setter
    def corner(self, value: object) -> None:
        self._corner = value

    @property
    def column_headers(self) -> list[object]:
        """The column headers (row 1, from B1)."""
        return list(self._column_headers)

    @property
    def row_labels(self) -> list[object]:
        """The row labels (column A, from A2)."""
        return list(self._row_labels)

    @property
    def data(self) -> list[list[object]]:
        """The data region (B2 onward) as a list of rows."""
        return [list(row) for row in self._data]

    # ---------------------------------------------------------- label access

    def _row_index(self, label: object) -> int:
        try:
            return self._row_labels.index(label)
        except ValueError:
            raise KeyError(f"no row labeled {label!r}") from None

    def _column_index(self, header: object) -> int:
        try:
            return self._column_headers.index(header)
        except ValueError:
            raise KeyError(f"no column headed {header!r}") from None

    def at(self, row_label: object, column_header: object) -> object:
        """The single value at the intersection of *row_label* and *column_header*."""
        return self._data[self._row_index(row_label)][self._column_index(column_header)]

    def row(self, row_label: object) -> list[object]:
        """The data row for *row_label* (labels not included)."""
        return list(self._data[self._row_index(row_label)])

    def column(self, column_header: object) -> list[object]:
        """The data column for *column_header* (headers not included)."""
        j = self._column_index(column_header)
        return [row[j] for row in self._data]

    # ----------------------------------------------------- positional access

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

    def _resolve(
        self, ref: str | None, row: int | None, column: int | str | None
    ) -> tuple[int, int]:
        """Turn a ref like ``"B2"`` or a ``row``/``column`` pair into a 1-based (row, col)."""
        if ref is not None:
            if row is not None or column is not None:
                raise TypeError("give either ref or row=/column=, not both")
            return coordinate_to_tuple(ref)
        if row is None or column is None:
            raise TypeError("give a ref like 'B2', or both row= and column=")
        col_num = column_index_from_string(column) if isinstance(column, str) else column
        return row, col_num

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

    def cell(
        self,
        ref: str | None = None,
        *,
        row: int | None = None,
        column: int | str | None = None,
    ) -> object:
        """A single value by Excel coordinate.

        Give either a reference like ``"B2"`` or ``row=`` and ``column=``
        (1-based; ``column`` may be a letter or a number). Row 1 is the header
        row and column 1 is the label column, so ``cell(row=2, column=2)`` is
        the first data cell.
        """
        r, c = self._resolve(ref, row, column)
        kind, i, j = self._classify(r, c)
        if kind == "corner":
            return self._corner
        if kind == "header":
            return self._column_headers[j]
        if kind == "label":
            return self._row_labels[i]
        return self._data[i][j]

    # -------------------------------------------------------------- mutation

    def set(self, row_label: object, column_header: object, value: object) -> None:
        """Set the single data value at *row_label* / *column_header*."""
        i = self._row_index(row_label)
        j = self._column_index(column_header)
        self._data[i][j] = value

    def set_cell(
        self,
        ref: str | None = None,
        *,
        row: int | None = None,
        column: int | str | None = None,
        value: object,
    ) -> None:
        """Set a value by Excel coordinate.

        A cell in row 1 sets a column header; row 1 / column 1 sets the corner;
        column 1 sets a row label; anything else sets data.
        """
        r, c = self._resolve(ref, row, column)
        kind, i, j = self._classify(r, c)
        if kind == "corner":
            self._corner = value
        elif kind == "header":
            self._column_headers[j] = value
        elif kind == "label":
            self._row_labels[i] = value
        else:
            self._data[i][j] = value

    def set_row(self, row_label: object, values: Iterable[object]) -> None:
        """Replace the data row for *row_label*; ``len(values)`` must match the column count."""
        new_row = list(values)
        i = self._row_index(row_label)
        expected = self._width()
        if len(new_row) != expected:
            raise ValueError(f"expected {expected} values, got {len(new_row)}")
        self._data[i] = new_row

    def set_column(self, column_header: object, values: Iterable[object]) -> None:
        """Replace the data column for *column_header*; ``len(values)`` must match the row count."""
        new_col = list(values)
        j = self._column_index(column_header)
        if len(new_col) != len(self._data):
            raise ValueError(f"expected {len(self._data)} values, got {len(new_col)}")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row[j] = value

    def add_row(self, label: object, values: Iterable[object]) -> None:
        """Append a labeled data row; ``len(values)`` must match the column count."""
        new_row = list(values)
        if self._data and not self._row_labels:
            raise ValueError("this table has no row labels; use the raw layer to add rows")
        if (self._data or self._column_headers) and len(new_row) != self._width():
            raise ValueError(f"expected {self._width()} values, got {len(new_row)}")
        self._data.append(new_row)
        self._row_labels.append(label)

    def add_column(self, header: object, values: Iterable[object]) -> None:
        """Append a labeled data column; ``len(values)`` must match the row count."""
        new_col = list(values)
        if any(self._data) and not self._column_headers:
            raise ValueError("this table has no column headers; use the raw layer to add columns")
        if len(new_col) != len(self._data):
            raise ValueError(f"expected {len(self._data)} values, got {len(new_col)}")
        for data_row, value in zip(self._data, new_col, strict=True):
            data_row.append(value)
        self._column_headers.append(header)

    def drop_row(self, label: object) -> None:
        """Remove the row labeled *label*."""
        i = self._row_index(label)
        del self._data[i]
        del self._row_labels[i]

    def drop_column(self, header: object) -> None:
        """Remove the column headed *header*."""
        j = self._column_index(header)
        del self._column_headers[j]
        for data_row in self._data:
            del data_row[j]

    def rename_row(self, old: object, new: object) -> None:
        """Change a row label."""
        self._row_labels[self._row_index(old)] = new

    def rename_column(self, old: object, new: object) -> None:
        """Change a column header."""
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

    def __len__(self) -> int:
        return len(self._data)

    def __iter__(self) -> Iterator[list[object]]:
        return iter([list(row) for row in self._data])

    def __getitem__(self, row_label: object) -> list[object]:
        return self.row(row_label)

    def __contains__(self, row_label: object) -> bool:
        return row_label in self._row_labels

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
