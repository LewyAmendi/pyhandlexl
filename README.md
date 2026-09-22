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

> **Stable.** pyhandlexl is at 1.0 — the public API won't change incompatibly
> within the 1.x series (see
> [Stability and versioning](#stability-and-versioning)), and every change is
> recorded in the
> [changelog](https://github.com/LewyAmendi/pyhandlexl/blob/main/CHANGELOG.md).

## Install

```bash
pip install pyhandlexl
```

Requires Python 3.10+ and [openpyxl](https://openpyxl.readthedocs.io/) 3.1.3 or newer
(below 4), which pip installs for you. Tested on Linux, Windows and macOS, on Python
3.10 to 3.14, against both the oldest and the newest openpyxl it allows. The package
is fully typed (`py.typed`) and passes `mypy --strict`.

Optional extras, none of which the library needs: `pyhandlexl[pandas]` and
`pyhandlexl[numpy]` for [analysis tools](#using-your-data-with-pandas-numpy-and-matplotlib),
and `pyhandlexl[fast]` (lxml), which makes writes about a quarter faster.

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
The table remembers what it read, so a later `write()` can merge with changes
another writer made in between — see
[Two people editing the same table](#two-people-editing-the-same-table).

Row labels, column headers, and the corner are always `str`. A numeric header
cell (`2024`) is read back as `"2024"`; building a `Table` by hand with a
non-`str` label/header/corner raises `TypeError`. Data values keep their type.

### The whole table at once

```python
t.set_corner("name")  # changes the value of cell A1
t.data                 # a TableData snapshot of everything, described just below
t.row_labels           # ['Alice', 'Bob']  — the labels alone, a cheap copy
t.column_headers       # ['q1', 'q2']      — the headers alone, a cheap copy
t.corner               # the corner cell's value (str)
```

`t.data` copies every value each time you touch it, so for the labels, the headers or the
corner use the three properties above — always in a loop, where `t.data.row_labels` on every
pass would be quadratic (4 ms a time on a table of 4,000 rows).

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
Table.from_dict(table, corner="", *, name, style=None, column_types=None)
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

The `values` you give `add_row`, `insert_row`, `set_row`, `set_column`,
`add_column`, and `insert_column` must be a sequence: a bare string raises
`TypeError` rather than being quietly split into one cell per character
(`add_row("x", "ab")` is a mistake, not two cells). The same goes for the rows you
hand to `write_sheet`/`append_rows` and the `pyhandlexl.grid` functions.

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
already exist for either call — see [Files](#files). If another writer changed
the table since you read it, `write` merges their changes with yours instead of
overwriting them — see the next section.

### Two people editing the same table

More than one script — or person — can work on the same table without losing
each other's changes, as long as they aren't writing at the very same instant.
A `Table` remembers what it last saw on disk. If someone else changed the
table since, `write()` reads what is on disk *now* and **merges** it with your
edits, rather than replacing it with your stale copy:

```python
alice = Table.read("log.xlsx", "Log")
bob   = Table.read("log.xlsx", "Log")

alice.add_row("alice-1", [1, 2])
bob.add_row("bob-1", [3, 4])

alice.write("log.xlsx")
bob.write("log.xlsx")        # merges: the file now has BOTH rows

bob.row_labels               # [..., 'alice-1', 'bob-1'] — bob's object is updated too
bob.last_merge               # MergeReport(rows_added=('alice-1',), ...)
```

Rows and columns are matched by label/header, and the merge is decided per
row, column, and cell:

| You and the other writer… | Result |
|---|---|
| added different rows/columns | both appear (yours after theirs when you both appended) |
| edited different cells | both edits survive |
| added the same label with the same values | one row, no conflict |
| deleted a row/column the other left alone | it's gone |
| renamed a row/column the other edited | your new name, their edit |
| changed the corner, style, or a column type (only one side did) | that change is kept |
| **changed the very same thing differently** | **yours wins**, with a warning |
| deleted a row/column the other **edited** | your delete wins, with a warning |
| edited a row/column the other **deleted** | your edit restores it, with a warning |

Where yours wins over something the other writer did, `write()` still succeeds
and emits a `MergeConflictWarning` naming each place it overwrote (up to ten,
then a count). It isn't an error — use `warnings.filterwarnings("error",
category=MergeConflictWarning)` if you'd rather stop instead. After any
merge, `t.last_merge` is a `MergeReport` with `rows_added`, `rows_removed`,
`columns_added`, `columns_removed`, `cells_updated` (your cells that took the
other writer's value), and `conflicts`; it is `None` if nothing had changed on
disk. Values are compared the way Excel stores them, so `5` vs `5.0`, or a
`date` vs its midnight `datetime`, is not mistaken for an edit — while `True`
vs `1` is.

Details worth knowing:

- Column types are checked **after** merging, so a restriction the other writer
  added can reject a value you set: `ColumnTypeError`, nothing is written, and
  your `Table` object is left exactly as it was.
- A `Table` that was built in memory and never read or created has nothing to
  merge against, so `write()` simply replaces the table, as before.
- Merging needs unique labels. A table with duplicate row labels or column
  headers (only possible from a hand-edited file) can't be matched row by row,
  so yours replaces the disk version, with a warning.

**What this does not do: it is not a lock.** The merge closes the gap between
your `read` and your `write` — the long part, where you're editing. It does
*not* stop two writes that land at the very same instant: each checks for
changes and then saves as separate steps, so a write that sneaks in between them
can still be overwritten. Writes to *different* tables in one workbook can also
overwrite each other for the same reason (every write re-saves the whole
workbook). If your writers can genuinely collide, serialise them yourself
(a queue, a single writer process, or a lock of your own around the write).
pyhandlexl deliberately does not take file locks itself, because OS file
locking behaves differently across platforms and filesystems.

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

A write repaints only what changed. If the style and the table's shape are the
same as on disk, nothing is repainted — only the values are rewritten — and adding
or dropping rows or columns repaints just the old and the new last row and column
(the only cells whose border changes; banding follows a row's position, not its
data). So formatting you added by hand *inside* a table, such as a highlighted cell
or a currency format, survives an ordinary write. Changing `t.style` repaints the whole
table, and a number format follows the value's type: a date is always written as a date.

Styling never touches `t.data` or equality — `t1 == t2` compares data,
headers, labels, corner, and column types only, however differently the two
are styled. It has no effect on `.csv` files — there's nothing there to style.

### Restricting a column's type

A column defaults to accepting any type `check_cell_value` allows — restrict
one to a single Excel-native type instead, and a value that doesn't match
raises `ColumnTypeError` the moment you try to set it:

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

A value can be a type Excel can store at all, and still fail this because
it isn't the type *this* column was restricted to.

**Checked immediately, not on the next write.** `set_column_type` checks the
column's *existing* data right away, and refuses (leaving the restriction
unapplied) if a value already there doesn't fit. Every method that sets or adds
data — `set_cell`, `set_row`, `set_column`, `add_row`/`insert_row` — checks the
value it's given against the target column's restriction the same way, so a
value that doesn't fit is rejected the moment you try to set it, not on the
next `create()`/`write()`; nothing changes when it's rejected. Two things are
still only checked at `write()`, because neither goes through an edit at all:
a restriction that arrives from another writer, through a
[merge](#two-people-editing-the-same-table), and a value that was already
there — say, a cell someone edited in Excel after the restriction was set —
when the table was [read](#reading).

Renaming, inserting, or dropping a column moves or drops its restriction along
with it, and a newly inserted column always starts as `ColumnType.ANY`. If the
reserved schema sheet is
deleted and rebuilt from markers (see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet)),
restrictions are **inferred** from what each column currently holds rather
than restored — nothing in a cell records that a column was restricted.
That can go either way: an empty or mixed column comes back `ColumnType.ANY`
(a restriction lost), and a column you'd deliberately left unrestricted, but
which happens to hold a single type, comes back restricted to it (one
invented). Check `t.column_types` after a rebuild and `set_column_type` back
anything that's wrong.

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

The dates are persisted in the reserved schema sheet, and unlike size and
sheet they can't be reconstructed from the table itself: if that sheet is
deleted and rebuilt from markers, nothing in a cell records when a table was
created or last modified — and, unlike column types, there's no data to
infer them from — so a rebuilt table reports `created_at`/`modified_at` as
`None` rather than a guessed time. Size and sheet still come back exact,
same as always.

### Equality

```python
t1 == t2   # compares data, headers, labels, corner, and column types
```

A `Table` has no `len()` or `in` of its own, so the call site says what's being checked:
`len(t.row_labels)` (or `t.info.n_rows`) for the row count, `"Bob" in t.row_labels` for
membership.

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
only relocates a table it already has a schema entry for. Instead, every
operation that reads the schema (`Table.read`, `Table.write`, `Table.create`,
`table_info`, `list_tables`, `delete_table`, `sheet_kind`, `write_sheet` (to a
sheet that already exists), `append_rows`, `clear_all_sheet_data`,
`delete_sheet`, `rename_sheet`) **rebuilds it** the moment it finds it missing:
it scans every sheet for `"TABLE NAME"` markers, emits a
`SchemaRebuiltWarning`, and — once the call
succeeds — saves the rebuilt sheet (re-marked with the same tab color and
comment) so the scan isn't repeated next time. The warning names the tables
found and points at the line of *your* code that made the call. A workbook
with no tables has nothing to rebuild, so it neither warns nor gains a
schema sheet.

**What a rebuild restores:**

- **Which tables exist, and their names** — every marker found, on any sheet.
- **Each table's exact position and size** — sheet, top-left corner, rows and
  columns. Markers bound each other directly, and there's nowhere legitimate
  for real content to sit past a table's true edge.
- **All of the table's contents** — data, row labels, column headers, and
  corner. Those live in the cells, never in the schema, so they can't be lost
  by deleting it.
- **A style and column types, as a best-effort guess** (next list).

**What a rebuild does not — and cannot — restore:**

- **Creation and modification times.** Nothing in a cell records them, and
  there's no data to infer them from, so `t.info.created_at`/`.modified_at`
  come back `None` rather than a guessed time. (See [Table metadata](#table-metadata).)
- **The exact style.** It is read back from the cells themselves and can be
  imperfect: a table with exactly one data row, for instance, can never have
  its row-banding detected (there's no second row to compare against), so it
  always comes back reporting no banding even if it originally had some.
- **Your declared column types.** They are *inferred* from the values each
  column currently holds — a column whose values all share one Excel-native
  type comes back restricted to it, and an empty or mixed column comes back
  `ColumnType.ANY`. That can go either way: a restriction can be **lost**, or
  one **invented** for a column you'd left unrestricted. See
  [Restricting a column's type](#restricting-a-columns-type).
- **A table whose marker is gone.** If the `"TABLE NAME"` cell was deleted or
  overwritten, there is nothing to find: the table is not rebuilt, and
  reading it raises `TableNotFoundError`.
- **A table name claimed by two markers.** It can't be safely resolved, so that
  table is left out of the rebuilt schema (named in the warning) rather than
  guessing which one is real; it isn't reachable by name until you resolve it
  by hand. Every unambiguous table is unaffected.

After a rebuild, check `t.style`, `t.column_types`, and `t.info` on the tables
you care about, set back anything that's wrong, and `write()` it.

**Workbooks written by older versions** keep working: a reserved sheet with
fewer columns than today's (written before column types, or before creation and
modification times, existed) is read with sensible defaults — the default
style, unrestricted columns, `None` dates — and upgraded to the current layout
by the next write.

**Editing a table by hand.** Numeric and date headers or labels are read as text.
A blank data cell is `None`. Duplicate labels or headers are tolerated on read.
A *blank* header or label can't be read as a table (`MalformedTableError`), and the error says which cell
(`... the column header in cell C2 is blank ...`) so you can fix it in Excel. Rows or
columns inserted above or left of a table are healed; deleting the marker cell (or a
column through it) makes the table unfindable. Error values (`#DIV/0!`) read as their
text, and so do formulas — with a `FormulaWarning` (see
[Formulas are not supported](#formulas-are-not-supported)).

A table may be any width Excel allows (16,383 data columns), and its column types
are always recorded in full — for very wide tables in a compact form, since a cell
holds at most 32,767 characters.

**One limit of the rebuild:** a cell holding the words `TABLE NAME` is only mistaken
for a table's marker when it sits at the very top of a sheet with text beside it — for
example a *grid* sheet whose first row starts `TABLE NAME | hello`. Inside a real
table (a label or data value that says it) it is recognised for what it is.

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

```
--------+----+-----
| name  | q1 | q2 |
--------+----+-----
| Alice | 10 | 20 |
| Bob   | 30 | 40 |
--------+----+-----
```

The default (`head=5, tail=5`) shows everything with no divider if the table
has 10 rows or fewer — truncation only kicks in past that. `rows=` overrides
the head/tail defaults outright. Prints a bordered, column-aligned grid to the
console **and returns that same text** — a debug convenience, unrelated to
cell formatting in the `.xlsx` (still out of scope; see [Not in
scope](#not-in-scope)). A value's own newlines are shown as `\n` rather than
left as real line breaks, which would otherwise split a row across lines.
`rows`, `head`, and `tail` must each be a non-negative `int` or `None` — a
negative value raises `ValueError` and a non-`int` raises `TypeError`, rather
than silently doing something confusing with Python's slice semantics.

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
  number, date, or boolean. That is a promise, not a stopgap: a CSV stays
  all text. `read_sheet`'s trimming/`pad` behavior still applies, but pads
  with `""` instead of `None`.
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

Prints a bordered, column-aligned block to the console and returns that same
text — a debug convenience, with the same truncation rules (and the same
`rows`/`head`/`tail` validation, and newline handling) as
[`Table.show`](#displaying-a-table), minus the header row (a grid has no
headers to show). Ragged rows are padded with `""` for display only; the grid
itself is untouched.

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
number, date, or boolean — and a field that starts with `=` is text, like every other
field. `encoding` defaults to `"utf-8-sig"`, which also
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

## Using your data with pandas, numpy and matplotlib

Neither library is needed to use pyhandlexl, and it never imports either one until you
call something below. Install what you use:

```bash
pip install "pyhandlexl[pandas]"     # DataFrames (brings numpy too)
pip install "pyhandlexl[numpy]"      # arrays only
```

### A table as a DataFrame

```python
from pyhandlexl import Table

t = Table.read("budget.xlsx", "Budget")
df = t.to_dataframe()          # row labels -> index, headers -> columns, corner -> index name
df["q1"].mean()
df.plot(kind="bar")            # matplotlib, through pandas
```

`to_dataframe` returns a copy, so changing it doesn't change the table. A blank cell is
`NaN` (`NaT` in a date or duration column). Each column's dtype is inferred from what it
holds, and a column restricted to a [`ColumnType`](#restricting-a-columns-type) is settled
to match: `DATE` as `datetime64`, `DURATION` as `timedelta64`, `BOOLEAN` as `bool` (or
pandas' nullable `boolean` when the column has blanks), `NUMBER` as `int64` or `float64`.
That last one follows Excel, which has one kind of number: a column of whole numbers comes
back `int64` even if it went in as `float64`.

### A DataFrame as a table

```python
result = df.assign(total=df["q1"] + df["q2"])
Table.from_dataframe(result, name="Result").create("budget.xlsx", sheet="Analysis")
```

The index becomes the row labels and its name the corner; the columns become the headers.
**Both must be strings**, since a table's labels and headers are text — a `DataFrame` with
the default integer index is refused, with the fix in the message:

```python
df.index = df.index.astype(str)        # keep the numbers, as text
df = df.set_index("id")                # or make one of the columns the labels
```

A `MultiIndex` is refused too (flatten it first with `reset_index()`). Values are stored as
the plain Python values they hold, and every missing value — `NaN`, `NaT`, `pd.NA` — becomes
a blank cell. `style=` and `column_types=` work as in the constructor, and
`infer_column_types=True` restricts each column to the type its values have (numbers,
booleans, dates, durations, times, text — a column of mixed or no values stays
unrestricted), so a later edit that doesn't fit it is refused with `ColumnTypeError`.

`Table.from_dataframe(...).write(path)` replaces a table that already exists. It builds a new
in-memory table that has never read the file, so there is nothing to merge with: it
overwrites.

### A table as a numpy array

```python
arr = t.to_numpy()                       # the data, without the labels
labels, headers = t.row_labels, t.column_headers
```

With no `dtype` the array gets the natural one for what the table holds: `int64`, `float64`,
`bool`, `datetime64[us]` or `timedelta64[us]` when every cell is of that kind (a blank is
`nan` or `NaT`), and `object` for text or a mixture (a blank stays `None`). Pass
`dtype=float` (or any other) to choose; a blank can't be an integer, and the message says
so.

### Grids

For a plain sheet — no row labels, no headers you name — `pyhandlexl.grid` has the same two
conversions:

```python
from pyhandlexl import read_sheet, write_sheet, grid

df = grid.to_dataframe(read_sheet("data.xlsx"))            # first row -> column names
write_sheet("out.xlsx", grid.from_dataframe(df))           # column names -> first row
write_sheet("out.xlsx", grid.from_dataframe(df, index=True))   # index -> first column
```

`header=False` treats every row as data. `write_sheet(path, df)` would not do: iterating a
DataFrame yields its column names, not its rows.

### Anywhere a value is written

numpy and pandas values are accepted wherever a value is written — `set_cell`, `add_row`,
`write_sheet`, `append_rows`, a `Table` built from an array — and stored as the plain value
they hold: `np.int64(3)` as `3`, `np.bool_` as a boolean, `np.datetime64` and
`pd.Timestamp` as a `datetime`, `np.timedelta64` and `pd.Timedelta` as a `timedelta`.
**`nan`, `pd.NaT` and `pd.NA` become a blank cell** (`inf` is still refused: Excel can't
store it). `check_cell_value` accepts the same.

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

A read-only file is never overwritten: a write to one raises `FileReadOnlyError`
(a kind of `FileLockedError`) before anything happens — on every platform, not just
Windows. A save keeps the file's permissions, and if the path is a symlink the file it
points to is the one updated.

`create_workbook` only makes **`.xlsx`** files (`ValueError` for any other name, so a
typo can't produce, say, an `.xlsx` inside `data.csv` that Excel refuses to open), and
the directory must already exist. Existing **`.xlsm`/`.xltm`** workbooks can be
read and written and keep their macros; an `.xltx` template stays a template.
Charts, merged cells, column widths, frozen panes, data validation, comments,
hyperlinks, formulas and defined names elsewhere in a workbook are preserved when
you write to it — but anything openpyxl itself can't keep (images without
Pillow installed, form controls, slicers) is not.

**Formulas keep their formula, but not their calculated result.** Excel stores each
formula's last result next to it; openpyxl can't keep those, so every write drops them
for the whole workbook. Excel and LibreOffice recalculate when they open the file, so
they show the numbers again; a tool that only reads the file (pandas, a preview pane,
`openpyxl.load_workbook(data_only=True)`) sees blanks until Excel has opened and
re-saved it.

## Sheet management

```python
from pyhandlexl import (
    list_sheets, sheet_exists, create_sheet, delete_sheet, rename_sheet, list_tables,
    sheet_kind, clear_all_sheet_data,
)

list_sheets(path)                 # ['Sheet', 'Data']
sheet_exists(path, "Data")        # True
create_sheet(path, "Results")     # SheetExistsError if it already exists
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

Excel treats sheet names as **case-insensitive**, so `create_sheet`,
`rename_sheet`, `write_sheet`, and `append_rows` refuse a name that differs from
an existing sheet's only by capitalisation (`SheetExistsError`) instead of letting the
file quietly end up with `data1` when you asked for `data`. (Changing just a
sheet's own capitalisation — `rename_sheet(path, "log", "LOG")` — is fine.) The
same goes for the reserved schema sheet's name, `_pyhandlexl_tables`, in *any*
capitalisation, or with stray spaces around it: no call can create, rename to,
write to, or otherwise use it (`SheetKindError`). A sheet name also can't contain a
control character or a character XML can't hold, or begin or end with an
apostrophe (`SheetNameError`).

## Safe writes

Every write goes through the same steps:

1. Refuse if the file is read-only (`FileReadOnlyError`), and follow it if it is a
   symlink so the real file is the one updated.
2. Save to a temporary file in the same directory.
3. Verify it is a readable `.xlsx` **and that every XML part in it parses** — a
   file that merely opens as a zip is not enough.
4. Give it the original's permissions and atomically replace the original
   (`os.replace`).

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
| `datetime`/`time`/`timedelta` with microseconds | rounded to the millisecond | Excel keeps time to a millisecond |
| `"a\r\nb"` or a lone `"\r"` in a data cell | `"a\nb"` | line breaks are stored as `\n`, the same on every platform |
| `"=1+1"` (a `str` starting with `=`) | `"=1+1"`, stored as text | pyhandlexl has no formulas: see below |
| `nan`, `float("nan")`, `pd.NaT`, `pd.NA` | `None` | Excel has no such values; every kind of missing value is an empty cell |
| `np.int64(3)`, `np.bool_`, `pd.Timestamp`, … | `3`, `True`, `datetime`, … | numpy and pandas values are stored as the plain Python value they hold |

Rejected outright (`CellTypeError`), because a silent change or loss is worse than
an error: `Decimal` (would silently become `float`), timezone-aware
`datetime`/`time` (Excel has no timezone), and any non-cell type (`list`,
`dict`, `bytes`, `complex`, …). Also refused, since Excel can't hold them:

- a string containing a character an `.xlsx` file cannot contain — control
  characters such as NUL, lone surrogates, `U+FFFE`/`U+FFFF`. A single one of
  these makes the *whole workbook* unreadable, not just its cell;
- a string longer than **32,767 characters**, which Excel would silently cut short;
- the infinities (`nan` is fine: it is stored as an empty cell);
- an `int` too large to be a float;
- a date before **1900-01-01** or after 9999-12-31 23:59:59.999;
- a carriage return in a table's **name, row label, column header, or corner** —
  those are found by their exact text, and Excel stores line breaks as `\n`
  (a data cell just has its line breaks normalised, as the table above says).

A `str` that looks numeric (`"007"`) stays a `str` in both directions.

### Formulas are not supported

pyhandlexl never writes a formula. A string is text, even one that starts with `=`:
`"=1+1"` is stored as the text `=1+1` (with Excel's invisible quote prefix, so it stays
text if someone edits the cell) and reads back as `"=1+1"`. That holds everywhere —
data cells, grids, row labels, headers, table names, imported CSV fields.

A workbook that already has formulas is another matter. Reading a table or a sheet that
holds them gives you each formula's *text* (`"=A1+B1"`, not its result) and raises a
`FormulaWarning`, because **writing that data back stores the text in place of the
formula.** Formulas in ranges you never read and write, and on other sheets, are left as
they are (though a write drops their calculated results — see [Files](#files)). If you
need formulas, use openpyxl directly.

`check_cell_value(value)` runs this check on a single value if you want to
validate before writing. Three more helpers are public for the same reason:

```python
from pyhandlexl import check_cell_value, check_sheet_name, check_dimensions, is_valid_xlsx

check_cell_value(value)         # CellTypeError if Excel can't store it faithfully
check_sheet_name("My sheet")    # SheetNameError if it isn't a legal worksheet name
check_dimensions(n_rows, n_cols)  # DimensionError if it won't fit on an Excel sheet
is_valid_xlsx(path)             # True if the file opens as a workbook (never raises)
```

## Errors

There are two kinds, and it is worth knowing which is which:

- **A problem with the data or the file** raises one of the library's own exceptions, all of
  which derive from `PyhandlexlError` — so `except PyhandlexlError` catches every one of them
  (below). Each is also the standard exception it resembles, so `except ValueError` or
  `except KeyError` keeps working too.
- **A wrong call** — a value of the wrong type, an unknown row or column label, a duplicate
  label, a position out of range, a wrong-length row — raises the standard `TypeError`,
  `ValueError`, `KeyError` or `IndexError`, as it would from any Python function. The
  file-system errors, `FileNotFoundError` and `FileExistsError`, are also the standard ones.
  These are mistakes in the code, not failures to handle.

| Exception | Also a | Meaning |
|---|---|---|
| `SheetNameError` | `ValueError` | invalid worksheet name (not a string, empty, too long, illegal characters) |
| `SheetExistsError` | `SheetNameError` | a sheet with that name — or one differing only by capitalisation — already exists |
| `DimensionError` | `ValueError` | data exceeds Excel's 1,048,576 × 16,384 grid |
| `SheetNotFoundError` | `KeyError` | no worksheet with that name |
| `FileLockedError` | `OSError` | file stayed locked (open in Excel) through every retry |
| `FileReadOnlyError` | `FileLockedError` | file is read-only, so nothing was written — retrying can't help; catching `FileLockedError` still covers it |
| `CellTypeError` | `TypeError` | a value is not something Excel can store faithfully (a foreign type, an XML-illegal or over-long string, `inf`, an out-of-range date, …) |
| `ColumnTypeError` | `TypeError` | a value doesn't match its column's `ColumnType` restriction |
| `TableNotFoundError` | `KeyError` | no named table with that name, or its marker is gone |
| `TableExistsError` | `ValueError` | a named table with that name already exists |
| `MalformedTableError` | `ValueError` | a table can't be read because a header or row label was left blank — the message names the cell |
| `SheetKindError` | `ValueError` | the sheet already holds the other kind of data (table vs. grid), or its name is the reserved schema sheet's (in any capitalisation) |
| `InvalidFileError` | `ValueError` | the file is not a readable workbook (or CSV): not a zip, damaged, or a part that won't parse |

An idempotent "make sure this sheet exists" is just
`try: create_sheet(path, "Log")` / `except SheetExistsError: pass`.

`SchemaRebuiltWarning`, `MergeConflictWarning` and `FormulaWarning` are not in this table on
purpose — they're `Warning`s (via Python's `warnings` module), not
`PyhandlexlError`s. The operation that triggers them still succeeds; see
[Multiple named tables on one sheet](#multiple-named-tables-on-one-sheet) for
when the first fires, [Two people editing the same table](#two-people-editing-the-same-table)
for the second, and [Formulas are not supported](#formulas-are-not-supported) for the third.

## Limitations

What to know before relying on it, so none of it is a surprise:

- **It is not a database.** There are no transactions and no locks. The
  [merge on write](#two-people-editing-the-same-table) protects two people who edit at
  different times; two saves landing in the same instant can still lose one of them.
  Serialise writers yourself if that can happen.
- **Every write loads and re-saves the whole workbook,** so its cost follows the size of
  the *file*, not of your edit — writing back a single changed cell of a table costs about
  the same as writing the whole table fresh; there is no partial write. Repainting only
  what changed, rather than the whole table, cuts that by about 16%, but the save itself
  still scales with the size of the file, not the size of the edit — a small workbook
  writes in a small fraction of the time a large one does.
- **A read only parses the sheets it needs.** Reading a table or a sheet, listing tables
  or sheets, or asking a sheet's kind opens the workbook read-only, so on a workbook of
  several big sheets, reading one small table — or listing what's in it — commonly costs
  well under 2% of loading the whole file (40 to over 100 times faster, depending on how
  many sheets there are and what you ask for). Reading a *big* sheet itself gains much
  less: that read is bound by how fast the underlying parser gets through the cells, not
  by how much of the rest of the workbook you skipped. It is meant for thousands of rows,
  not hundreds of thousands — and for a plain grid, `append_rows` costs only what the new
  rows cost. Installing [lxml](https://lxml.de/) (`pip install "pyhandlexl[fast]"`) makes
  writes about a quarter faster; reads are unaffected.
- **The whole table is held in memory,** as Python objects.
- **Excel must not have the file open on Windows** — a write retries for a moment, then
  raises `FileLockedError`. On macOS and Linux nothing stops a save from replacing a file
  that is open in Excel, so Excel's copy silently goes stale.
- **Only Excel's data model.** Cells hold what Excel can store — text up to 32,767
  characters, numbers, booleans, and dates and times without a time zone, to the
  millisecond. See the [round-trip notes](#round-trip-notes) for the small changes that
  forces. Formulas are not supported (see above), and neither are `Decimal` or a time zone.
- **A sheet holds tables or a grid, not both,** and tables sit side by side starting at
  row 1 with one empty column between them; you can't place one at an arbitrary cell.
- **A write drops the calculated results of any formulas in the workbook** — the
  formulas stay, and Excel or LibreOffice recalculate them on opening (see [Files](#files)).
- **A `Table` object is not thread-safe.**
- **Files are checked by reading them back with openpyxl,** not by opening them in Excel
  or LibreOffice. The library refuses everything it knows would make Excel complain (an
  over-long sheet name, an illegal character, a value Excel would cut short); if Excel
  ever does complain about a file it wrote, that is a bug worth reporting.
- **`.xlsx`, `.xlsm`, `.xltx`, `.xltm` and `.csv` only** — no `.xls`, no password-protected
  workbooks.

## Not in scope

`pyhandlexl` deliberately does **not** handle: arbitrary cell-level
formatting (`Table`'s own [styling](#styling-a-table) is a curated set of
choices, not a general "format any cell" API), conditional formatting,
formulas, charts, images, merged cells, `.xls` (old format), or password
protection / encryption. For any of that, use openpyxl directly. (Charts,
merged cells and the like that are already in a workbook are left alone when
you write to it — see [Files](#files) — but pyhandlexl has no way to create or
edit them.)

## Stability and versioning

pyhandlexl follows [semantic versioning](https://semver.org/). From 1.0.0 the **public API** does
not change incompatibly within a major version: a script that works on 1.0 works on every
1.x release. The public API is:

- every name importable from the top-level package — `import pyhandlexl` and everything in
  `pyhandlexl.__all__` — and the functions of `pyhandlexl.grid`, with their parameters (by name
  and position), what they return, and which exceptions they raise;
- the file format: a workbook written by any 1.x release opens in any other, and the reserved
  schema sheet is always readable by later versions.

New features arrive in minor releases and only add: new functions, new keyword arguments with
defaults, new exception subclasses (which are always subclasses of the exception they refine, so
existing `except` clauses keep working).

Not part of the API, and free to change in any release: the internal modules (anything whose
name starts with an underscore, and the submodules `pyhandlexl.core`, `.table`, `.validate` and
the like when imported directly instead of from the top level), the exact wording of messages
and warnings, the `repr` of any object, and performance. The validators `check_cell_value`,
`check_sheet_name` and `is_valid_xlsx` take their argument by position, so it has no public
name.

A feature is not removed without a deprecation warning that lasts at least one minor release;
removal happens only in the next major version.

## Development

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\Activate.ps1
pip install -e ".[dev]"
pytest
ruff check . && ruff format --check .
mypy
```

## License

MIT — see [LICENSE](https://github.com/LewyAmendi/pyhandlexl/blob/main/LICENSE).
