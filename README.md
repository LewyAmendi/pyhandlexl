# pyhandlexl

[![CI](https://github.com/LewyAmendi/pyhandlexl/actions/workflows/ci.yml/badge.svg)](https://github.com/LewyAmendi/pyhandlexl/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/pyhandlexl)](https://pypi.org/project/pyhandlexl/)

**Use an Excel file as a lightweight database for a small project.**

For internal tools, prototypes, and anything where people need to read or edit
the data themselves, a spreadsheet is often enough — portable, familiar, no
server to run. The friction is in the code: [openpyxl](https://openpyxl.readthedocs.io/)
gives you full control of the workbook, which means working in worksheets, cell
coordinates, ranges, and dimensions.

`pyhandlexl` is a simpler layer over openpyxl for the common case — read
organised data from a sheet, change it, write it back. It treats a worksheet as
a structured table, so your code can focus on what the data *represents* rather
than where it sits.

```python
from pyhandlexl import Table

t = Table.read("budget.xlsx", "Budget")
t.read_cell(row="Alice", column="q2")   # 20  — by name, and typed
t.add_row("Carol", [1, 2])
t.write("budget.xlsx")                   # safe, atomic write
```

`pyhandlexl` splits cleanly in two, by how organised your data is:

- **Organised data → [`Table`](#organised-data-the-table-class).** Your sheet
  has column headers and row labels — a real table. Address everything by
  name (`t.read_cell(row="Revenue", column="North")`), not by cell position.
- **Unorganised data → [the grid layout](#unorganised-data-the-grid-layout).**
  Plain grids, exports, odd one-off dumps — no headers or labels to name.
  `read_sheet`/`write_sheet` give you a `list[list]`; `pyhandlexl.grid` edits
  it by position.

> **Under active development.** Usable today — expect new capabilities with each
> release, and some API changes before it stabilises. See the
> [changelog](CHANGELOG.md).

## Install

```bash
pip install pyhandlexl
```

Requires Python 3.10+.

## Quickstart

`pyhandlexl` never creates a file or a table implicitly, so a full lifecycle
looks like this:

```python
from pyhandlexl import Table, create_workbook, create_sheet

create_workbook("budget.xlsx")
create_sheet("budget.xlsx", "Data")

Table(
    data=[[10, 20], [30, 40]],
    column_headers=["q1", "q2"],
    row_labels=["Alice", "Bob"],
    name="Budget",
).create("budget.xlsx", sheet="Data")   # first placement — explicit
```

From then on, read and write it by name — no need to remember which sheet
it's on:

```python
t = Table.read("budget.xlsx", "Budget")

t.read_cell(row="Alice", column="q2")   # 20     (int)
t.read_row("Bob")                       # [30, 40]
t.read_column("q1")                     # [10, 30]

t.set_cell(row="Alice", column="q1", value=99)   # edit in place
t.add_row("Carol", [1, 2])
t.write("budget.xlsx")     # one safe, atomic write
```

> **Values keep their type.** `read_sheet` and `Table` return each cell as its
> native Python type — `str`, `int`, `float`, `bool`, `datetime`, `date`,
> `time`, `timedelta` — and an empty cell as `None`. Writing preserves type
> too; a value that isn't one of those raises `CellTypeError`. **Row labels,
> column headers, and the corner are the exception** — they are always coerced
> to `str`, since you address rows and columns by name. See
> [round-trip notes](#round-trip-notes) for the small type changes Excel forces.

## Organised data: the `Table` class

**Use this when your data has structure — column headers and row labels
that name what's in each cell.** A `Table` treats a worksheet region as
exactly that: a header row, a label column, a corner, and data — and it
stays in sync as you add, drop, or rename rows and columns. Every `Table`
has a **name**, which is how you find it again; a worksheet can hold more
than one, side by side — see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet).

If your sheet is just a plain grid with nothing to name, skip ahead to
[Unorganised data: the grid layout](#unorganised-data-the-grid-layout) instead.

### Reading

```python
Table.read(path, name)
```

Finds the table called `name` anywhere in the workbook — no need to know
which sheet it's on. `TableNotFoundError` if there's no table by that name.

Row labels, column headers, and the corner are always `str`. A numeric header
cell (`2024`) is read back as `"2024"`; building a `Table` by hand with a
non-`str` label/header/corner raises `TypeError`. Data values keep their type.

### The whole table at once

```python
t.corner       # value of cell A1 (settable: t.corner = "name")
t.data         # a TableData snapshot
```

```python
d = t.data
d.rows              # [[10, 20], [30, 40]]   (B2 onward, by row — typed)
d.columns           # [[10, 30], [20, 40]]   (same data, by column)
d.row_labels        # ['Alice', 'Bob']       (column A, from A2 — str)
d.column_headers    # ['q1', 'q2']           (row 1, from B1 — str)
d.corner            # value of cell A1       (str)
```

Every field is a fresh copy — mutating `t.data.rows` does not change the table.

```python
t.to_dicts()   # [{'q1': 10, 'q2': 20}, {'q1': 30, 'q2': 40}]  — one dict per data row
```

`to_dicts` keys each data row by column header; row labels aren't included —
use `t.data` alongside it if you need those too. Build a table the other way
with `Table.from_dicts`:

```python
from pyhandlexl import Table

t = Table.from_dicts(
    [{"q1": 10, "q2": 20}, {"q1": 30, "q2": 40}],
    row_labels=["Alice", "Bob"],
    name="Budget",
)
```

Every dict must have the same keys, in the same order — that order becomes
the column headers (`ValueError` otherwise).

### Access by label

```python
t.read_row("Bob")            # [30, 40]   — a data row (no label), typed
t.read_column("q1")          # [10, 30]   — a data column (no header)
```

Unknown labels raise `KeyError`. `add_row` / `add_column` / `rename_row` /
`rename_column` refuse to create a duplicate label (`ValueError`); a table read
from a file may still contain duplicates, in which case reads take the first
match.

### `read_cell` / `set_cell` — a single value, by position or by label

Both take the same addressing: a ref like `"B2"`, or `row=`/`column=` as a
matching pair — **both ints** for a 1-based Excel position (row 1 is the
header row, column 1 is the label column), or **both strings** for a row
label / column header pair. Mixing types raises `TypeError`.

```python
t.read_cell("B2")                       # 10 — first data cell, by position
t.read_cell(row=2, column=2)            # 10 — same thing, spelled out
t.read_cell(row="Alice", column="q1")   # 10 — same value, by label

t.read_cell(row=1, column=2)   # 'q1'    — a column header (str)
t.read_cell(row=2, column=1)   # 'Alice' — a row label (str)
t.read_cell("A1")               # the corner (str)
```

By position, `read_cell` can reach *any* cell — header, label, corner, or
data. By label it always reads data, wherever that row/column intersection
actually lives.

A string given through `row=`/`column=` is always a label lookup — it does
**not** accept a column letter like `"B"` for a position. Use a plain number
(`column=2`) or `ref="B2"` for letter-based positions.

### Editing (in place, returns `None`)

```python
t.set_cell("B2", value=99)                       # by position — ref
t.set_cell(row=2, column=2, value=99)            # by position — row=/column= as ints
t.set_cell(row="Alice", column="q1", value=99)   # by label    — row=/column= as strings
t.set_row("Bob", [50, 60])          # replace a row   (length must match)
t.set_column("q1", [1, 2])          # replace a column (length must match)

t.add_row("Carol", [1, 2])          # append a labelled row
t.add_column("q3", [5, 6])          # append a labelled column

t.insert_row(1, "Carol", [1, 2])       # insert as the new first row
t.insert_column(2, "q0", [5, 6])       # insert as the new 2nd column

t.drop_row("Bob")
t.drop_column("q2")

t.rename_row("Alice", "ALICE")
t.rename_column("q1", "Q1")
t.corner = "name"
```

`set_cell` only ever touches **data** — addressing a header, row label, or the
corner by position raises `ValueError`; use `rename_row`, `rename_column`, or
`t.corner = value` for those. Wrong-length values, and a name that would
duplicate an existing label/header, raise `ValueError`; unknown labels raise
`KeyError`; a non-`str` row label, column header, or corner raises `TypeError`.
Data values may be any type; a value Excel can't store is caught on `.write()`
(`CellTypeError`), not when it's set.

`insert_row`/`insert_column` take a 1-based position among existing data rows/
columns: `1` inserts as the new first one; `len(...) + 1` inserts as the last
— the same result as `add_row`/`add_column`, which are exactly that special
case. Same constraints as `add_row`/`add_column` otherwise; an out-of-range
position raises `IndexError`.

You can also build a table up from nothing before placing it — `create()` is
only needed once, for that first placement:

```python
from pyhandlexl import create_workbook, Table

t = Table([], column_headers=["q1", "q2"], name="Budget")
t.add_row("Alice", [10, 20])
create_workbook("new.xlsx")
t.create("new.xlsx", sheet="Sheet")
```

### Writing

```python
t.create(path, sheet)   # first placement on a sheet
t.write(path)            # every write after that
```

`create` places a brand-new named table on `sheet`: `TableExistsError` if the
name is already taken, `SheetNotFoundError` if the sheet doesn't exist. Once a
table has been created, `write` reassembles its headers, labels, and data and
writes it back to its tracked location — no `sheet=` needed, and
`TableNotFoundError` if the table was never created (or has since been
deleted). The file must already exist for either call — see
[Files](#files).

### Equality

```python
t1 == t2            # compares data, headers, labels, corner
```

Row count and row-label membership go through `t.data` instead of `len()`/`in`,
so the call site says what's being checked: `len(t.data.rows)`,
`"Bob" in t.data.row_labels`.

### Multiple named tables on one sheet

A worksheet isn't limited to one table — a "Sales" table and an "Inventory"
table can live side by side. Tables stack **left to right** with exactly one
empty column between them, and always start at row 1.

```python
from pyhandlexl import Table, create_workbook, create_sheet

create_workbook("shop.xlsx")
create_sheet("shop.xlsx", "Data")

sales = Table(
    data=[[100, 200], [150, 250]],
    column_headers=["North", "South"],
    row_labels=["Q1", "Q2"],
    name="Sales",
)
sales.create("shop.xlsx", sheet="Data")          # first placement — explicit

inventory = Table(data=[[10], [20]], column_headers=["Units"], row_labels=["A", "B"], name="Inventory")
inventory.create("shop.xlsx", sheet="Data")       # lands to the right of Sales

sales = Table.read("shop.xlsx", "Sales")          # no sheet= needed — the name finds it
sales.add_column("East", [300, 350])
sales.write("shop.xlsx")                          # Inventory shifts right automatically
```

- **`list_tables(path)`** — every named table in the workbook, across all sheets.
- **`delete_table(path, name)`** — removes a named table; `TableNotFoundError`
  if it doesn't exist. Leaves the space empty — other tables on the sheet are
  not shifted to close the gap, and the name is free to reuse afterwards.
- Table names are **unique per workbook**, not per sheet.
- **Growing a table's columns** (`add_column`, or writing back wider data)
  automatically shifts every table to its right, on the same sheet, further
  right to make room — this can rewrite more than one table's position in a
  single `.write()`. **Growing rows never shifts anything**, since nothing sits
  below a table.
- **Duplicate row labels/column headers within one table** are blocked the
  same way regardless of how many tables share the sheet.

**How it stays correct if you edit the sheet by hand.** Every named table's
first cell literally contains the text `"TABLE NAME"`, with the table's name
in the cell beside it. Before trusting its tracked position, `pyhandlexl`
checks that marker is still there. If someone has inserted or deleted columns
by hand and it's moved, `pyhandlexl` scans that sheet, finds it by its marker,
and repairs the tracked position — this can make `Table.read` write to the
file even though it looks like a pure read. If the marker is gone entirely
(the table was deleted, or the name cell was overwritten), you get
`TableNotFoundError` rather than silently wrong data. This self-heal only
recovers a *moved* table on its *original* sheet — if a table was deleted, cut
to a different sheet, or resized by hand, `pyhandlexl` reports it missing
rather than guessing further.

The tracking data itself lives in a reserved worksheet, `_pyhandlexl_tables`
— it shows up in `list_sheets()` like any other sheet. Leave it alone.

### Displaying a table

```python
t.show()                     # default: first 5 and last 5 rows, "..." between
t.show(rows=10)               # only the first 10 data rows
t.show(head=2, tail=2)        # first 2 and last 2 rows
t.show(head=None, tail=None)  # every row, no truncation
```

The default (`head=5, tail=5`) shows everything with no divider if the table
has 10 rows or fewer — truncation only kicks in past that. `rows=` overrides
the head/tail defaults outright. Prints a plain, aligned, whitespace-padded
grid to the console — a debug convenience, unrelated to cell formatting in the
`.xlsx` (still out of scope; see [Not in scope](#not-in-scope)).

## Unorganised data: the grid layout

**Use this when your data has no headers or labels to name — a plain grid,
an export, an odd one-off layout.** There's nothing here to address by name,
so everything works by position: a `list[list]` you read, edit by row/column
index, and write back. If your sheet *does* have column headers and row
labels, use [`Table`](#organised-data-the-table-class) instead — it'll save
you from re-inventing header/label handling by hand.

```python
from pyhandlexl import read_sheet, write_sheet, append_rows

read_sheet(path, sheet=None, *, pad=False)
```

Returns `list[list[object]]` — each cell as its native type (`str`, `int`,
`float`, `bool`, `datetime`, `date`, `time`, `timedelta`), an empty cell as
`None`. Trailing `None` values are trimmed from each row (a fully empty row
becomes `[]`); `pad=True` right-pads every row with `None` to the widest row's
length instead.

```python
write_sheet(path, rows, sheet=None, *, orientation="rows")
```

Replaces the target sheet with `rows` (other sheets untouched); adds `sheet` if
it does not exist. Values are written with their type preserved — no conversion;
`None` leaves the cell empty. A value that isn't a type Excel can store raises
`CellTypeError`. `orientation="columns"` writes each inner list *down a column*
instead of across a row.

```python
append_rows(path, rows, sheet=None)
```

Appends after the last row. Empty input is a no-op.

### Editing a grid: `pyhandlexl.grid`

Helpers for the `list[list[object]]` that `read_sheet` returns. Each takes a
grid and returns a **new** grid, so they compose in a pipeline. Rows and
columns are **1-based** (row 1 is the first row), matching `Table`.

```python
from pyhandlexl import read_sheet, write_sheet, grid

g = read_sheet("data.xlsx")

g = grid.set_value(g, 2, 3, "changed")   # one cell
g = grid.set_row(g, 1, ["a", "b", "c"])  # replace a row
g = grid.set_column(g, 2, [1, 2, 3])     # replace a column
g = grid.insert_row(g, 2, [...])         # insert before row 2
g = grid.insert_column(g, 1, [...])      # insert as the new first column
g = grid.append_row(g, [...])
g = grid.append_column(g, [...])
g = grid.delete_row(g, 3)
g = grid.delete_column(g, 2)

g = grid.transpose(g)   # rows <-> columns (ragged rows padded with None)
g = grid.pad(g)         # rectangularise

write_sheet("data.xlsx", g)
```

`get_row(g, i)` / `get_column(g, j)` read a single line. Column operations that
take `values` require `len(values)` to equal the row count; out-of-range
indices raise `IndexError`. Gaps introduced by any of these (padding a short
row, `pad`, `transpose`) are filled with `None`.

## Files

Applies whether you're working with organised or unorganised data.
`pyhandlexl` **never creates a file implicitly** — this is deliberate, so a
typo in a path can't silently produce a stray workbook.

```python
create_workbook(path, *, sheet="Sheet")   # FileExistsError if the path is taken
delete_workbook(path)                      # FileNotFoundError if it isn't there
```

`create_workbook` makes a new empty `.xlsx` with one worksheet. `delete_workbook`
removes a workbook file, retrying while it is locked (open in Excel) before
raising `FileLockedError`, and refuses a path that isn't an Excel extension.

Every write operation — `write_sheet`, `append_rows`, `create_sheet`,
`Table.write` — raises `FileNotFoundError` if the file does not exist yet.

## Sheet management

```python
from pyhandlexl import (
    list_sheets, sheet_exists, create_sheet, delete_sheet, rename_sheet, list_tables,
)

list_sheets(path)                 # ['Sheet', 'Data']
sheet_exists(path, "Data")        # True
create_sheet(path, "Results")     # ValueError if it already exists
delete_sheet(path, "Old")         # refuses to delete the last sheet
rename_sheet(path, "Old", "New")
list_tables(path)                 # every named table in the workbook (all sheets)
```

`create_sheet` needs an existing file (`create_workbook` first). Sheet names are
validated everywhere: max 31 characters, none of `\ / ? * [ ] :`, and
`"History"` is reserved by Excel.

## Safe writes

Every write goes through the same steps:

1. Save to a temporary file in the same directory.
2. Verify it is a readable `.xlsx`.
3. Atomically replace the original (`os.replace`).

If any step fails the temporary file is removed and the original is left exactly
as it was. If the target is locked (open in Excel), writes retry briefly before
raising `FileLockedError`.

## Round-trip notes

Excel forces a few small type changes on `write` → `read`:

| You write | You read back | Why |
|---|---|---|
| `5.0` (float) | `5` (int) | Excel stores all numbers as float; openpyxl returns `int` when the value is whole |
| `date(2026, 1, 1)` | `datetime(2026, 1, 1, 0, 0)` | Excel has no date-only type |
| `""` (empty string) | `None` | Excel doesn't distinguish an empty string from a blank cell |
| int larger than 2⁵³ | loses precision | float limit |

Rejected outright (`CellTypeError`): `Decimal` (would silently become `float`),
timezone-aware `datetime`/`time` (Excel has no timezone), and any non-cell type
(`list`, `dict`, `bytes`, `complex`, …). A `str` that looks numeric (`"007"`)
stays a `str` in both directions.

`check_cell_value(value)` runs this check on a single value if you want to
validate before writing.

## Errors

All raised exceptions derive from `PyhandlexlError`:

| Exception | Also a | Meaning |
|---|---|---|
| `SheetNameError` | `ValueError` | invalid worksheet name |
| `DimensionError` | `ValueError` | data exceeds Excel's 1,048,576 × 16,384 grid |
| `SheetNotFoundError` | `KeyError` | no worksheet with that name |
| `FileLockedError` | `OSError` | file stayed locked through every retry |
| `CellTypeError` | `TypeError` | a value is not a type Excel can store |
| `TableNotFoundError` | `KeyError` | no named table with that name, or its marker is gone |
| `TableExistsError` | `ValueError` | a named table with that name already exists |
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

MIT — see [LICENSE](https://github.com/LewyAmendi/pyhandlexl/blob/main/LICENSE).
