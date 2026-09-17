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

Have a `.csv` file instead of `.xlsx`? `read_sheet`/`write_sheet`/
`append_rows`/`grid` work directly on one too, as a raw grid — see
[CSV rules](#csv-rules-for-read_sheet-write_sheet-and-append_rows). To move
data *between* CSV and `.xlsx` instead, see
[Moving data between CSV and .xlsx](#moving-data-between-csv-and-xlsx).

> **Under active development.** Usable today — expect new capabilities with each
> release, and some API changes before it stabilises. See the
> [changelog](https://github.com/LewyAmendi/pyhandlexl/blob/main/CHANGELOG.md).

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
t.set_corner("name")  # changes the value of cell A1
t.data                 # a TableData snapshot — read the corner via t.data.corner
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
t.to_dict()
# {'Alice': {'q1': 10, 'q2': 20}, 'Bob': {'q1': 30, 'q2': 40}}
```

`to_dict` keys the outer dict by row label and the inner one by column
header, so both axes survive in a single structure. Build a table the other
way with `Table.from_dict`:

```python
Table.from_dict(table, corner="", *, name)
```

```python
from pyhandlexl import Table

t = Table.from_dict(
    {"Alice": {"q1": 10, "q2": 20}, "Bob": {"q1": 30, "q2": 40}},
    name="Budget",
)
```

The outer keys become the row labels, in order. Every inner dict must have
the same keys, in the same order — that order becomes the column headers
(`ValueError` otherwise); `table` itself, and every row in it, must be a
mapping (`TypeError` otherwise). `table` must not be empty — with no rows
there's nothing to infer the column headers from. A duplicate row label or
column header (only possible on a table read from a file) collapses to its
last value on `to_dict`, since dict keys must be unique.

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
(`column=2`) or `ref="B2"` for letter-based positions. A malformed `ref`
(`""`, `"not a ref"`) raises `ValueError`.

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
t.set_corner("name")
```

`set_cell` only ever touches **data** — addressing a header, row label, or the
corner by position raises `ValueError`; use `rename_row`, `rename_column`, or
`set_corner` for those. Wrong-length values, and a name that would duplicate
an existing label/header, raise `ValueError`; unknown labels raise `KeyError`;
a non-`str` row label, column header, or corner raises `TypeError`. A row
label or column header can never be `""` — that raises `ValueError` too (the
corner has no such restriction; it may be empty). Data values may be any
type; a value Excel can't store is caught on `.write()` (`CellTypeError`),
not when it's set. `drop_column` refuses to remove a table's only remaining
column (`ValueError`) — `column_headers` is required and can never end up
empty.

`insert_row`/`insert_column` take a 1-based position among existing data rows/
columns: `1` inserts as the new first one; `len(...) + 1` inserts as the last
— the same result as `add_row`/`add_column`, which are exactly that special
case. Same constraints as `add_row`/`add_column` otherwise; a non-`int`
position raises `TypeError`, an out-of-range one raises `IndexError`.

You can also build a table up from nothing before placing it — `create()` is
only needed once, for that first placement. `column_headers=` and `name=` are
the only required arguments; data and row labels can start empty and grow
with `add_row`:

```python
from pyhandlexl import create_workbook, Table

t = Table(column_headers=["q1", "q2"], name="Budget")
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
name is already taken, `SheetNotFoundError` if the sheet doesn't exist,
`SheetKindError` if `sheet` already holds plain grid data (see
[One kind of data per sheet](#one-kind-of-data-per-sheet)). Once a table has
been created, `write` reassembles its headers, labels, and data and writes it
back to its tracked location — no `sheet=` needed, and `TableNotFoundError` if
the table was never created (or has since been deleted). The file must
already exist for either call — see [Files](#files).

### Styling a table

Every table is visually styled when it's created or written: bold, filled
headers and row labels; alternating row colors; a thick border around the
whole table with a heavier line separating headers/labels from data. This
is a real property of the table, not a one-time paint job — adding a row or
column extends the same look to cover it, and a table shifted right to make
room for a growing neighbor (see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet))
keeps its look at the new position too.

```python
from pyhandlexl import Table, TableStyle

t = Table(data=[[10, 20]], column_headers=["q1", "q2"], row_labels=["Alice"], name="Budget")
t.style                       # TableStyle.DEFAULT, unless you passed style=
t.style = TableStyle.MINIMAL  # takes effect on the next create()/write()
```

Three ready-made looks:

| | header/labels | data banding | border |
|---|---|---|---|
| `TableStyle.DEFAULT` | bold, black on olive green | light gray stripe | thick outer, medium header line |
| `TableStyle.MINIMAL` | bold, plain colors | none | thick outer, medium header line |
| `TableStyle.NONE` | plain | none | none |

For anything else, build one — every field is a plain 6-digit hex color
(`""` means "none," for `header_fill`/`band_fill`/`border_color` only —
font colors are always required):

```python
TableStyle(
    header_font_name="Calibri", header_font_size=11, header_font_color="000000",
    header_bold=True, header_fill="76933C",
    data_font_name="Calibri", data_font_size=11, data_font_color="000000",
    band_fill="F2F2F2", border_color="000000",
)
```

Styling never touches `t.data` or equality — `t1 == t2` compares data,
headers, labels, corner, and column types only, however differently the two
are styled. It has no effect on `.csv` files — there's nothing there to style.

### Restricting a column's type

A column defaults to accepting any type `check_cell_value` allows — restrict
one to a single Excel-native type instead, and a value that doesn't match
raises `ColumnTypeError` on the next `create()`/`write()`:

```python
from pyhandlexl import Table, ColumnType

t = Table(
    data=[[1000, "Alice"]],
    column_headers=["amount", "name"],
    row_labels=["r1"],
    name="Orders",
    column_types={"amount": ColumnType.NUMBER},
)
t.column_types              # {'amount': ColumnType.NUMBER, 'name': ColumnType.ANY}
t.set_column_type("name", ColumnType.TEXT)
```

`ColumnType.NUMBER`, `.TEXT`, `.BOOLEAN`, `.DATE`, `.TIME`, `.DURATION`, and
`.ANY` (the default) follow *Excel's* type model, not Python's exactly:
`NUMBER` covers both `int` and `float` (Excel stores every number as a
float and doesn't distinguish them — see
[Round-trip notes](#round-trip-notes)), and `DATE` covers both
`datetime.date` and `datetime.datetime` (Excel has no date-only type). A
blank cell (`None`) is always allowed regardless of a column's type — the
restriction governs what a real value may be, not whether the cell has been
filled in yet.

Checked the same time as `CellTypeError` — a value can be a type Excel can
store at all, and still fail this because it isn't the type *this* column
was restricted to. Restricting a column doesn't touch data already in it
until the next `create()`/`write()`; renaming, inserting, or dropping a
column moves or drops its restriction along with it, and a newly inserted
column always starts as `ColumnType.ANY`. Like style, a column's type
restriction can't be recovered if the reserved schema sheet is deleted and
rebuilt from markers (see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet))
— a rebuilt table always comes back reporting `ColumnType.ANY` for every
column.

### Table metadata

```python
t.info   # a TableInfo snapshot — not data, not style: when, where, how big
```

```python
i = t.info
i.created_at       # when create() first placed it (UTC datetime, or None)
i.modified_at      # when write()/create() last touched it (UTC, or None)
i.n_rows           # data row count
i.n_cols           # data column count
i.sheet            # which worksheet it's on (str, or None)
i.style            # same object as t.style
i.column_types     # same mapping as t.column_types
```

`sheet`, `created_at`, and `modified_at` are `None` until the table has
actually been placed with `create()` (or loaded with `read()`) — a table
you're still building up in memory has no sheet or history yet.
`modified_at` only reflects *this* table's own `create()`/`write()` calls —
a table shifted right to make room for a growing neighbour (see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet))
is not itself modified, so its `modified_at` is untouched by that.

For checking a table's metadata without reading (and type-converting) its
actual row data — say, listing every table's size and dates in a workbook
— there's a standalone lookup that skips all of that:

```python
from pyhandlexl import table_info

table_info(path, "Sales")   # same TableInfo, without an existing Table object
```

Both are persisted in the reserved schema sheet, and both are subject to
the same rebuild limitation as style and column types: if that sheet is
deleted and reconstructed from markers, there's nothing in a cell that
records when a table was created or last modified, so a rebuilt table
reports `created_at`/`modified_at` as `None` rather than a guessed time —
size and sheet still come back exact, same as always.

### Equality

```python
t1 == t2   # compares data, headers, labels, corner, and column types
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
- **`table_info(path, name)`** — a table's [metadata](#table-metadata) (dates,
  size, sheet, style, column types) without reading its row data.
- **`delete_table(path, name)`** — removes a named table; `TableNotFoundError`
  if it doesn't exist. Leaves the space empty — other tables on the sheet are
  not shifted to close the gap, and the name is free to reuse afterwards.
- Table names are **unique per workbook**, not per sheet.
- **Growing a table's columns** (`add_column`, or writing back wider data)
  automatically shifts every table to its right, on the same sheet, further
  right to make room — this can rewrite more than one table's position in a
  single `.write()`. **Growing rows never shifts anything**, since nothing sits
  below a table. A shifted table keeps its
  [style](#styling-a-table) at the new position.
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
— `list_sheets()` never shows it, since it isn't a sheet you created or can
write to. Leave it alone: it's marked with a red sheet tab and a warning
comment on its first cell, every time it's written, so it's hard to miss
even for someone opening the workbook by hand without having read this.

**If that reserved sheet is deleted entirely**, self-heal can't help — it
only relocates a table it already has a schema entry for. Instead,
`Table.read`, `Table.write`, `delete_table`, and `list_tables` all
automatically rebuild the whole schema by scanning every sheet for
`"TABLE NAME"` markers the moment they find it missing, emitting a
`SchemaRebuiltWarning` and persisting the reconstruction (re-marked with the
same tab color and comment) so the scan isn't repeated next time. A table's
position and size come back exact — markers bound each other directly, and
there's nowhere legitimate for real content to sit past a table's true
edge. Its **style is a best-effort reconstruction** read back from the
cells themselves, and can be imperfect: a table with exactly one data row,
for instance, can never have its row-banding detected (there's no second
row to compare against), so it always comes back reporting no banding even
if it originally had some. Its **column-type restrictions cannot be
recovered at all** — see
[Restricting a column's type](#restricting-a-columns-type) — every column
comes back as `ColumnType.ANY`. Its **creation and modification times
cannot be recovered either** — see [Table metadata](#table-metadata) —
`t.info.created_at`/`.modified_at` both come back `None` rather than a
guessed time. Two markers found claiming the same name can't be safely
resolved either — that table is left out of the rebuilt schema (named in
the warning) rather than guessing which one is real; every unambiguous
table is unaffected.

### One kind of data per sheet

A worksheet holds **either** named tables **or** plain [grid data](#unorganised-data-the-grid-layout),
never both — mixing them would let one silently corrupt the other (a grid
write landing across a table's marker, say). Whichever kind writes to a
sheet first claims it:

```python
from pyhandlexl import sheet_kind, clear_all_sheet_data

sheet_kind(path, "Data")   # "empty", "grid", or "table"
```

- A fresh or freshly-cleared sheet is `"empty"` — either kind can claim it next.
- `write_sheet`/`append_rows` claim it as `"grid"`; `Table.create` claims it as `"table"`.
- Writing the other kind to an already-claimed sheet raises `SheetKindError`
  rather than writing something that would corrupt what's already there:
  `write_sheet`/`append_rows` refuse a `"table"` sheet, and `Table.create`
  refuses a `"grid"` sheet. The reserved `_pyhandlexl_tables` schema sheet
  can't be targeted directly by any of them either, or by `sheet_kind`
  itself — it isn't a sheet with a kind of its own. Neither can
  `delete_sheet` or `rename_sheet` — deleting it directly, or renaming it
  away (or renaming some *other* sheet onto its reserved name), is refused
  (`SheetKindError`); it's only ever meant to disappear as a side effect of
  deleting every table that lives on it.
- **`clear_all_sheet_data(path, sheet)`** wipes a sheet's cells, styles, and
  any tables tracked on it, back to `"empty"` — the one way to reverse a
  claim and let the sheet be reused as the other kind.

`delete_sheet` and `rename_sheet` keep the schema in sync with reality:
deleting a sheet forgets any tables that lived on it (instead of leaving
stale, unreachable schema entries), and renaming one moves its tables'
tracked location along with it — they stay readable under their same names.

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
`.xlsx` (still out of scope; see [Not in scope](#not-in-scope)). `rows`,
`head`, and `tail` must each be a non-negative `int` or `None` — a negative
value raises `ValueError` and a non-`int` raises `TypeError`, rather than
silently doing something confusing with Python's slice semantics.

## Unorganised data: the grid layout

**Use this when your data has no headers or labels to name — a plain grid,
an export, an odd one-off layout.** There's nothing here to address by name,
so everything works by position: a `list[list]` you read, edit by row/column
index, and write back. If your sheet *does* have column headers and row
labels, use [`Table`](#organised-data-the-table-class) instead — it'll save
you from re-inventing header/label handling by hand.

**These same functions work directly on a `.csv` file too** — pass a path
ending in `.csv` instead of `.xlsx` and they read/write the file itself
rather than a worksheet inside a workbook. A `.csv` file is treated purely as
a raw grid, never as a table: no types, no sheets, no size limit — see
[CSV rules](#csv-rules-for-read_sheet-write_sheet-and-append_rows) below.
`pyhandlexl.grid`'s editing functions need no changes to work either way,
since they only ever touch the `list[list]` these return, never a file.

```python
from pyhandlexl import read_sheet, write_sheet, append_rows

read_sheet(path, sheet=None, *, pad=False, orientation="rows")
```

For an **.xlsx** file: returns `list[list[object]]` — each cell as its
native type (`str`, `int`, `float`, `bool`, `datetime`, `date`, `time`,
`timedelta`), an empty cell as `None`. Trailing `None` values are trimmed
from each row (a fully empty row becomes `[]`); `pad=True` right-pads every
row with `None` to the widest row's length instead. `orientation="columns"`
returns each worksheet column as an inner list instead — the transpose of
`"rows"`, with the same trimming/`pad` rules applied down each column.

```python
write_sheet(path, rows, sheet=None, *, orientation="rows")
```

For an **.xlsx** file: replaces the target sheet with `rows` (other sheets
untouched); adds `sheet` if it does not exist. Values are written with their
type preserved — no conversion; `None` leaves the cell empty. A value that
isn't a type Excel can store raises `CellTypeError`. `orientation="columns"`
writes each inner list *down a column* instead of across a row. `sheet`
must not already hold table data (`SheetKindError`) — see
[One kind of data per sheet](#one-kind-of-data-per-sheet).

```python
append_rows(path, rows, sheet=None)
```

For an **.xlsx** file: appends after the last row. Empty input is a no-op.
Same `SheetKindError` restriction as `write_sheet`.

### CSV rules for read_sheet, write_sheet, and append_rows

```python
from pyhandlexl import create_csv, read_sheet, write_sheet, append_rows, grid

create_csv("data.csv")                    # files are never created implicitly
write_sheet("data.csv", [["a", "b"], [1, 2]])
g = read_sheet("data.csv")                # [['a', 'b'], ['1', '2']]

g = grid.insert_row(g, 1, ["h1", "h2"])   # the grid toolkit works unchanged
write_sheet("data.csv", g)

append_rows("data.csv", [["x", "y"]])
```

- **Every value becomes/comes back as a plain `str`** (`None` becomes `""`
  on write) — a CSV field has no other type, so nothing is inferred as a
  number, date, or boolean. May change in a future release; for now it's
  deliberately literal. `read_sheet`'s trimming/`pad` behavior still
  applies, but pads with `""` instead of `None`.
- **`sheet` must be `None`** — a `.csv` file has no sheets. Passing anything
  else raises `ValueError`.
- **No size limit** — `DimensionError`/`CellTypeError` never apply to a
  `.csv` write; it isn't bound by the `.xlsx` grid.
- **The file must already exist**, exactly like `.xlsx` — call `create_csv`
  first. `write_sheet` replaces the whole file's content (there's no sheet
  to isolate a change to) and writes atomically, the same as every other
  write in this library.

**Why `append_rows` is a separate function, not just
`write_sheet(path, grid.append_row(read_sheet(path), values))`:**
performance. That composition would read *every* existing row into Python
first, just to add a few more at the end — expensive once a sheet or file
is large. `append_rows` never does: it costs only what the *new* rows cost,
regardless of how much is already there. The two formats earn that
differently, because they can differently:

- For **.xlsx**, it loads the workbook (unavoidable — that's how you open
  one at all) but calls the equivalent of "append one row" only for the new
  rows; the existing ones are never walked or turned into Python values.
  Still written atomically, like every other `.xlsx` write.
- For **.csv**, it does better still: unlike a zip-based `.xlsx`, a plain
  text file can be modified without rewriting it, so `append_rows` opens it
  in append mode and writes only the new rows, touching none of the
  existing bytes. That's cheaper than the `.xlsx` case, but it does mean a
  `.csv` append isn't wrapped in the temp-file-then-replace safety
  `write_sheet` gets: a crash mid-write could leave a malformed trailing
  row, but — unlike a failed *replace* — can never lose or corrupt a byte
  that was already there.

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

```python
grid.show(g)                     # default: first 5 and last 5 rows, "..." between
grid.show(g, rows=10)             # only the first 10 rows
grid.show(g, head=2, tail=2)      # first 2 and last 2 rows
grid.show(g, head=None, tail=None)  # every row, no truncation
```

Prints a plain, aligned, whitespace-padded block to the console — a debug
convenience, with the same truncation rules (and the same `rows`/`head`/`tail`
validation) as [`Table.show`](#displaying-a-table). Ragged rows are padded
with `""` for display only; the grid itself is untouched.

## Moving data between CSV and .xlsx

There are two different things "CSV" might mean here, and two different
tools for them:

- **Editing a `.csv` file as itself** — `read_sheet`/`write_sheet`/
  `append_rows`/`grid` all work directly on a `.csv` path now, treating it
  as a raw grid (never a table). See
  [CSV rules](#csv-rules-for-read_sheet-write_sheet-and-append_rows) above.
  Nothing ever crosses into `.xlsx` here.
- **Moving data *between* a `.csv` file and an `.xlsx` worksheet** —
  `import_csv_to_xl`/`export_xl_to_csv`, below. These are one-shot
  **conversions**, not a live link and not editing in place.

```python
from pyhandlexl import import_csv_to_xl, export_xl_to_csv

import_csv_to_xl(csv_path, path, *, sheet=None, encoding="utf-8-sig")
export_xl_to_csv(path, csv_path, *, sheet=None, encoding="utf-8")
```

```python
import_csv_to_xl("results.csv", "experiments.xlsx", sheet="Log")
export_xl_to_csv("experiments.xlsx", "backup.csv", sheet="Log")
```

**`import_csv_to_xl`** reads `csv_path` and writes it into `sheet` of the
existing `.xlsx` file at `path` — same target semantics as `write_sheet`
(replaced if `sheet` already exists, created if it doesn't). Every field
becomes a `str` cell: CSV has no other type, so nothing is inferred as a
number, date, or boolean. `encoding` defaults to `"utf-8-sig"`, which also
strips a leading byte-order mark transparently (common in CSVs saved by
Excel on Windows).

**`export_xl_to_csv`** reads `sheet` from `path` and writes a brand-new CSV
file at `csv_path` — refuses to overwrite an existing file there
(`FileExistsError`), the same caution as `create_workbook`. Every value is
stringified with `str()` (`None` becomes an empty field): a lossy, one-way
conversion — re-importing the result gives back text, not the original
types. Written atomically, like every other write in this library (see
[Safe writes](#safe-writes)).

Both require the `.xlsx` file to already exist (`FileNotFoundError`
otherwise — see [Files](#files)), and both check the path they were given —
`import_csv_to_xl`'s `path` must be a workbook and `csv_path` must be
`.csv`, `export_xl_to_csv` the other way around (`ValueError` otherwise).
`import_csv_to_xl` also raises `DimensionError` if the CSV has more rows or
columns than an `.xlsx` worksheet can hold, checked *before* anything is
written.

## Files

Applies whether you're working with organised or unorganised data.
`pyhandlexl` **never creates a file implicitly** — this is deliberate, so a
typo in a path can't silently produce a stray workbook.

```python
create_workbook(path, *, sheet="Sheet")   # FileExistsError if the path is taken
delete_workbook(path)                      # FileNotFoundError if it isn't there
create_csv(path)                           # the .csv equivalent of create_workbook
```

`create_workbook` makes a new empty `.xlsx` with one worksheet. `delete_workbook`
removes a workbook file, retrying while it is locked (open in Excel) before
raising `FileLockedError`, and refuses a path that isn't an Excel extension.
`create_csv` makes a new empty `.csv` file — same guarantee, same
`FileExistsError` if something's already there, so a typo can't silently
overwrite or produce a stray file either way.

Every write operation — `write_sheet`, `append_rows`, `create_sheet`,
`Table.create`, `Table.write` — raises `FileNotFoundError` if the file does
not exist yet. This applies to a `.csv` target exactly the same as `.xlsx`.

## Sheet management

```python
from pyhandlexl import (
    list_sheets, sheet_exists, create_sheet, delete_sheet, rename_sheet, list_tables,
    sheet_kind, clear_all_sheet_data,
)

list_sheets(path)                 # ['Sheet', 'Data']
sheet_exists(path, "Data")        # True
create_sheet(path, "Results")     # ValueError if it already exists
delete_sheet(path, "Old")         # refuses to delete the last sheet, or the schema sheet; forgets its tables too
rename_sheet(path, "Old", "New")  # moves its tables' tracked location along with it; refuses the schema sheet either way
list_tables(path)                 # every named table in the workbook (all sheets)
sheet_kind(path, "Data")          # "empty", "grid", or "table" — see below
clear_all_sheet_data(path, "Data")  # wipes it back to "empty"
```

`create_sheet` needs an existing file (`create_workbook` first). Sheet names are
validated everywhere: max 31 characters, none of `\ / ? * [ ] :`, and
`"History"` is reserved by Excel. See
[One kind of data per sheet](#one-kind-of-data-per-sheet) for `sheet_kind`
and `clear_all_sheet_data`.

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
| `ColumnTypeError` | `TypeError` | a value doesn't match its column's `ColumnType` restriction |
| `TableNotFoundError` | `KeyError` | no named table with that name, or its marker is gone |
| `TableExistsError` | `ValueError` | a named table with that name already exists |
| `SheetKindError` | `ValueError` | the sheet already holds the other kind of data (table vs. grid), or is the reserved schema sheet |
| `InvalidFileError` | — | file is missing or not a readable `.xlsx` |

`SchemaRebuiltWarning` is not in this table on purpose — it's a `Warning`
(via Python's `warnings` module), not a `PyhandlexlError`. The operation
that triggers it still succeeds; see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet)
for when it fires.

## Not in scope

`pyhandlexl` deliberately does **not** handle: arbitrary cell-level
formatting (`Table`'s own [styling](#styling-a-table) is a curated set of
choices, not a general "format any cell" API), conditional formatting,
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
