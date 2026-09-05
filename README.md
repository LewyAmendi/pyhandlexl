# pyhandlexl

[![CI](https://github.com/LewyAmendi/pyhandlexl/actions/workflows/ci.yml/badge.svg)](https://github.com/LewyAmendi/pyhandlexl/actions/workflows/ci.yml)

**Make an Excel file the database for your next project.**

A spreadsheet is the most portable data store there is: every machine opens it,
anyone can read or edit it without knowing a query language, it versions as a
single file, and there is no server to run. `pyhandlexl` makes driving one from
Python dependable — read and write raw cell values in a known order, with writes
that leave the file intact even when things go wrong.

Built on [openpyxl](https://openpyxl.readthedocs.io/) and built to grow.

- **`Table`** — the main way in. Row 1 holds your column headers, column A holds
  your row labels, and everything from `B2` on is data. Read it, edit it by name,
  write it back.
- **`read_sheet` / `write_sheet`** — direct grid access for sheets that aren't a
  labelled table.

> **Version 0.1.0.** Usable today and under active development — expect new
> capabilities with each release, and some API changes as it matures.

## Install

Not on PyPI yet. From source:

```bash
git clone https://github.com/LewyAmendi/pyhandlexl
cd pyhandlexl
pip install -e .
```

Requires Python 3.10+.

## Quickstart

Given `budget.xlsx`:

|        | q1 | q2 |
|--------|----|----|
| **Alice** | 10 | 20 |
| **Bob**   | 30 | 40 |

```python
from pyhandlexl import Table

t = Table.read("budget.xlsx")

t.at("Alice", "q2")        # '20'
t.row("Bob")               # ['30', '40']
t.column("q1")             # ['10', '30']

t.set("Alice", "q1", 99)   # edit in place
t.add_row("Carol", [1, 2])
t.write("budget.xlsx")     # one safe, atomic write
```

> **Values are always strings.** `read_sheet` and `Table` coerce every cell to
> `str` (empty cells become `""`). Convert to numbers yourself where you need to.

## The `Table` class

### Reading

```python
Table.read(path, sheet=None, *, column_headers=True, row_labels=True)
```

- `sheet=None` reads the active sheet; pass a name for a specific one.
- `column_headers=False` — row 1 is ordinary data, `column_headers` is empty.
- `row_labels=False` — column A is ordinary data, `row_labels` is empty.

### Properties

```python
t.corner            # value of cell A1
t.column_headers    # ['q1', 'q2']       (row 1, from B1)
t.row_labels        # ['Alice', 'Bob']   (column A, from A2)
t.data              # [['10', '20'], ['30', '40']]   (B2 onward)
```

All four return copies — mutating them does not change the table.

### Access by label

```python
t.at("Alice", "q2")     # single value
t.row("Bob")            # a data row (no label)
t.column("q1")          # a data column (no header)
```

Unknown labels raise `KeyError`. If a label appears twice, the first match wins.

### Access by position (Excel coordinates)

Positions are exactly what Excel shows: row 1 is the header row, column 1 (`A`)
is the label column, so `B2` is the first data cell.

```python
t.cell("B2")                    # '10'   — first data cell
t.cell(row=2, column=2)         # '10'   — same thing
t.cell(row=1, column=2)         # 'q1'   — a column header
t.cell(row=2, column=1)         # 'Alice' — a row label
t.cell("A1")                    # the corner
```

`column` accepts a letter (`"B"`) or a 1-based number (`2`).

### Editing (in place, returns `None`)

```python
t.set("Alice", "q1", 99)            # set one data value by label
t.set_cell("B2", value=99)          # set one cell by coordinate
t.set_row("Bob", [50, 60])          # replace a row   (length must match)
t.set_column("q1", [1, 2])          # replace a column (length must match)

t.add_row("Carol", [1, 2])          # append a labelled row
t.add_column("q3", [5, 6])          # append a labelled column

t.drop_row("Bob")
t.drop_column("q2")

t.rename_row("Alice", "ALICE")
t.rename_column("q1", "Q1")
t.corner = "name"
```

Wrong-length values raise `ValueError`; unknown labels raise `KeyError`.

You can also build a table from nothing:

```python
t = Table([], column_headers=["q1", "q2"])
t.add_row("Alice", [10, 20])
t.write("new.xlsx")
```

### Writing

```python
t.write(path, sheet=None)
```

Reassembles headers into row 1 and labels into column A, then writes the whole
sheet. Other sheets in the file are left untouched.

### Container behaviour

```python
len(t)              # number of data rows
list(t)             # the data rows
t["Alice"]          # same as t.row("Alice")
"Bob" in t          # checks row labels
t1 == t2            # compares data, headers, labels, corner
```

## The raw layer

For sheets that are not a labelled table — plain grids, exports, odd layouts.

```python
from pyhandlexl import read_sheet, write_sheet, append_rows

read_sheet(path, sheet=None, *, pad=False)
```

Returns `list[list[str]]`. Trailing empty cells are trimmed from each row (a
fully empty row becomes `[]`); `pad=True` right-pads every row to the widest
row's length instead.

```python
write_sheet(path, rows, sheet=None, *, orientation="rows")
```

Replaces the target sheet with `rows` (other sheets untouched), creating the
file and sheet if needed. Values are written as-is — `str` stays `str`, `int`
stays `int`, `None` leaves the cell empty; there is no string-to-number
conversion. `orientation="columns"` writes each inner list *down a column*
instead of across a row.

```python
append_rows(path, rows, sheet=None)
```

Appends after the last row. Empty input is a no-op.

## Sheet management

```python
from pyhandlexl import (
    list_sheets, sheet_exists, create_sheet, delete_sheet, rename_sheet,
)

list_sheets(path)                 # ['Sheet1', 'Data']
sheet_exists(path, "Data")        # True
create_sheet(path, "Results")     # ValueError if it already exists
delete_sheet(path, "Old")         # refuses to delete the last sheet
rename_sheet(path, "Old", "New")
```

Sheet names are validated everywhere: max 31 characters, none of `\ / ? * [ ] :`,
and `"History"` is reserved by Excel.

## Safe writes

Every write goes through the same steps:

1. Save to a temporary file in the same directory.
2. Verify it is a readable `.xlsx`.
3. Atomically replace the original (`os.replace`).

If any step fails the temporary file is removed and the original is left exactly
as it was. If the target is locked (open in Excel), writes retry briefly before
raising `FileLockedError`.

## Errors

All raised exceptions derive from `PyhandlexlError`:

| Exception | Also a | Meaning |
|---|---|---|
| `SheetNameError` | `ValueError` | invalid worksheet name |
| `DimensionError` | `ValueError` | data exceeds Excel's 1,048,576 × 16,384 grid |
| `SheetNotFoundError` | `KeyError` | no worksheet with that name |
| `FileLockedError` | `OSError` | file stayed locked through every retry |
| `InvalidFileError` | — | file is missing or not a readable `.xlsx` |

## Not in scope

`pyhandlexl` deliberately does **not** handle: cell formatting, styles, fonts,
formulas, charts, images, merged cells, `.xls` (old format), or password
protection / encryption. For any of that, use openpyxl directly.

## Development

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\Activate.ps1
pip install -e ".[dev]"
pytest
ruff check . && ruff format --check .
```

## License

MIT — see [LICENSE](LICENSE).
