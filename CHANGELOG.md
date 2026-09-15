# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.9.0] — 2026-09-15

### Added
- **Every table is now visually styled** — bold, filled headers and row
  labels; alternating data-row colors; a thick outer border with a medium
  line separating headers/labels from data. Applied on `create()`/`write()`
  and re-applied in full every time, so adding a row or column extends the
  same look automatically, and a table shifted right by a growing neighbor
  keeps its look at the new position.
  - `TableStyle` (new, exported) — a frozen dataclass of hex colors, font
    name/size, and bold/fill/border choices. `TableStyle.DEFAULT`,
    `.MINIMAL`, and `.NONE` cover the common cases; construct your own for
    anything else. Validates its fields (must be a real 6-digit hex color,
    a positive font size, etc.) at construction.
  - `Table(..., style=...)` — defaults to `TableStyle.DEFAULT`. Readable and
    settable via `t.style`; a change takes effect (and is persisted) on the
    next `create()`/`write()`.
  - Persisted per table in the reserved `_pyhandlexl_tables` schema sheet,
    so a table keeps exactly the look it was created with even if a
    preset's definition changes later, and `Table.read()` restores it.
  - Style is excluded from `Table` equality (`t1 == t2` still compares only
    data, headers, labels, and corner) and has no effect on `.csv` files.

- **Automatic recovery if the reserved `_pyhandlexl_tables` schema sheet is
  deleted entirely** — something self-heal couldn't do, since it only
  relocates a table it already has a schema entry for. `Table.read`,
  `Table.write`, `delete_table`, and `list_tables` now rebuild the whole
  schema by scanning every sheet for `"TABLE NAME"` markers the moment they
  find it missing, and persist the reconstruction so the scan isn't
  repeated on the next call.
  - `SchemaRebuiltWarning` (new, exported) — a `Warning`, not a
    `PyhandlexlError`; the triggering operation still succeeds. Fires only
    when the scan actually finds something to recover — a genuinely
    table-less workbook stays silent and gets no schema sheet written.
  - Position and size come back exact: markers bound each other directly on
    a shared sheet, and nothing legitimate is ever placed past a table's
    real content otherwise. A table's *style* is a best-effort
    reconstruction read back from its cells, which can be imperfect — most
    notably, a table with exactly one data row can never have its row
    banding detected (there's no second row to compare against), so it
    always comes back reporting none.
  - Two markers found claiming the same name can't be safely resolved —
    that table is left out of the rebuilt schema (named in the warning)
    rather than guessing which one is real; every unambiguous table on the
    same workbook is unaffected.

### Fixed
- **Deleting or shrinking a table could leave "ghost" formatted-but-empty
  cells behind**, inflating the worksheet's saved dimensions the same way
  Excel treats a styled blank cell as "in use." `clear_region` (used by
  `delete_table` and by `write()` when a table gets smaller) now fully
  resets a cleared cell's style, not just its value.

## [0.8.0] — 2026-09-14

### Added
- **`read_sheet`/`write_sheet`/`append_rows` now work directly on a `.csv`
  file**, not just `.xlsx` — pass a path ending in `.csv` and they read/write
  the file itself, treated purely as a raw grid, never a table:
  - Every value is a plain `str` on read, and `str()`-converted on write
    (`None` becomes `""`) — a CSV field has no other type, so nothing is
    inferred as a number, date, or boolean. May change in a future release.
  - `sheet` must be `None` for a `.csv` path (`ValueError` otherwise) — it
    has no sheets.
  - No size limit — `DimensionError`/`CellTypeError` never apply to a `.csv`
    write.
  - The file must already exist, exactly like `.xlsx` — `create_csv(path)`
    is the `.csv` equivalent of `create_workbook`. `write_sheet` replaces
    the whole file and writes atomically; `append_rows` opens the file in
    append mode and writes only the new rows — it never reads the existing
    content, so appending stays cheap no matter how large the file already
    is, at the cost of the temp-file-then-replace safety `write_sheet` gets
    (a crash mid-write can leave a malformed trailing row, but can never
    lose or corrupt existing content).
  - `pyhandlexl.grid`'s editing functions needed no changes — they only
    ever operate on the `list[list]` these return, never touch a file.
- `import_csv_to_xl(csv_path, path, *, sheet=None, encoding="utf-8-sig")`
  and `export_xl_to_csv(path, csv_path, *, sheet=None, encoding="utf-8")` —
  one-shot conversions *between* a CSV file and an `.xlsx` worksheet (not a
  live link), renamed from `import_csv`/`export_csv` for clarity now that
  `read_sheet`/`write_sheet` also handle `.csv` natively. Behavior
  otherwise unchanged: `import_csv_to_xl` requires the `.xlsx` file to
  exist and raises `DimensionError` if the CSV is too big for an `.xlsx`
  grid, checked before anything is written; `export_xl_to_csv` refuses to
  overwrite an existing file at `csv_path` and writes atomically. Both now
  also validate that `csv_path`/`path` actually have the extension their
  name promises (`ValueError` otherwise).

## [0.7.1] — 2026-09-13

### Fixed
- Documentation only — no code changes. The README was missing several
  parameters and error behaviors added in 0.7.0: `from_dict`'s `corner=`
  argument, `drop_column`'s last-column guard, `insert_row`/`insert_column`'s
  non-`int` position `TypeError`, `read_cell`/`set_cell`'s malformed-ref
  `ValueError`, `show()`'s `rows`/`head`/`tail` validation, and `Table.create`
  in the list of operations that raise `FileNotFoundError`.

## [0.7.0] — 2026-09-13

### Added
- `grid.show()` — print a grid to the console as a plain aligned block, with
  the same truncation rules as `Table.show()` (first/last 5 rows by default,
  `rows=`/`head=`/`tail=` to override). A console convenience only.

### Changed
- **BREAKING: `column_headers=` is now a required, keyword-only argument.**
  A table's columns are its schema — every `Table` needs them, even before
  its first row. `data` now defaults to empty, so building one up before
  placing it is just `Table(column_headers=[...], name="...")`. The
  positional-data-only "grid" construction mode (no headers, no labels) is
  retired — it could never round-trip through `create()`/`write()`/`read()`
  correctly anyway; use the grid layout for genuinely positional data.
  `drop_column` now refuses to remove a table's last remaining column.

### Fixed
- **Data corruption on `create()`/`write()` for a table with headers but no
  rows yet.** The corner cell was omitted from the persisted header row
  whenever `row_labels` happened to be empty, silently shifting every column
  header over by one on read-back. Building a table's columns before its
  first `add_row()` — the README's own "build a table from nothing" pattern —
  triggered this.
- `Table(...)` now raises `ValueError` if it's given data with row labels
  missing or the wrong length — that combination could never round-trip
  through `create()`/`write()`/`read()`.
- A bare `str` passed for `column_headers`, `row_labels`, or `data` no longer
  silently iterates into one column/row per character — `Table(...)` now
  raises `TypeError` instead (`column_headers="ab"` used to quietly become
  two headers, `"a"` and `"b"`).
- `read_cell`/`set_cell` with a malformed `ref` (`""`, `"not a ref"`) now
  raise a clear `ValueError` instead of leaking an internal
  `UnboundLocalError` or a raw `int()` parsing error.
- `insert_row`/`insert_column` with a non-`int` `position`, and every `grid`
  function taking a row/column number, now raise a clear `TypeError` instead
  of a raw comparison/`list.insert` error.
- `Table.from_dict` with a non-mapping argument, or a non-mapping value for
  one of its rows, now raises a clear `TypeError` instead of a raw
  `AttributeError`/`TypeError` from the internal `dict()` conversion.
- `Table.read`/`write` and `delete_table` now raise `TypeError` for a
  non-`str` `name` instead of silently reporting `TableNotFoundError`.
- `Table.show(rows=..., head=..., tail=...)` now rejects a negative or
  non-`int` value instead of silently producing a confusing result via
  Python's slice semantics (`rows=-1` used to mean "every row but the last").

## [0.6.0] — 2026-09-13

### Added
- `Table.to_dict()` — the table as `dict[str, dict[str, object]]`, outer key
  row label, inner key column header, so both axes round-trip in one
  structure.
- `Table.from_dict(table, corner="", *, name)` — the inverse: builds a
  `Table` from a dict of dicts. The outer keys become the row labels; every
  inner dict must have the same keys, in the same order (`ValueError`
  otherwise) — that order becomes the column headers.
- `Table.insert_row(position, label, values)` / `insert_column(position, header,
  values)` — insert a labeled row/column at a 1-based position among the
  existing ones, instead of only appending. `add_row`/`add_column` are now the
  `position = len(...) + 1` special case of these. An out-of-range position
  raises `IndexError`.

## [0.5.0] — 2026-09-12

### Changed
- **BREAKING: every `Table` now requires a `name`, and is tracked by name —
  not by sheet.** Whole-sheet `Table` usage is retired.
  - `Table(...)` now takes a required, keyword-only `name`.
  - A table is placed once with `t.create(path, sheet)` — `TableExistsError`
    if the name is taken, `SheetNotFoundError` if the sheet doesn't exist.
  - `Table.read(path, name)` — finds a table by name anywhere in the
    workbook; `sheet=` is gone, since the name is enough.
  - `t.write(path)` — writes a table back to its tracked location;
    `sheet=` is gone here too. `TableNotFoundError` if `.create()` was never
    called for it.
  - The `column_headers=False` / `row_labels=False` reading modes are gone —
    every table always has both.
  - Several named tables can share one worksheet — see
    [Multiple named tables on one sheet](README.md#multiple-named-tables-on-one-sheet).
    Tables stack left to right with one empty column between them, always
    starting at row 1. Growing a table's columns shifts every table to its
    right on the same sheet; growing rows never shifts anything.
  - `list_tables(path)` — every named table in the workbook.
  - `delete_table(path, name)` — removes a named table; leaves the space empty
    (no shifting) and frees the name for reuse.
  - Each table's first cell holds a literal `"TABLE NAME"` marker plus its
    name; a workbook-wide schema sheet (`_pyhandlexl_tables`, reserved) caches
    positions. Reads/writes verify the marker before trusting the cached
    position and self-heal (and may write, even on a "read") if a table has
    moved; `TableNotFoundError` if it can't be found at all.
  - Table names are unique per workbook. Duplicate row labels/column headers
    within one table are blocked the same as before.

### Added
- `Table.show(*, rows=None, head=5, tail=5)` — print a table to the console as
  a plain aligned grid. Defaults to the first and last 5 rows (everything, if
  10 rows or fewer); pass `head=None, tail=None` for every row. A debug
  convenience, unrelated to cell formatting in the workbook.
- `TableNotFoundError(PyhandlexlError, KeyError)` and
  `TableExistsError(PyhandlexlError, ValueError)`.

## [0.4.0] — 2026-09-10

### Changed
- **Cell values keep their type.** `read_sheet` and `Table` now return each
  cell as its native Python type (`str`, `int`, `float`, `bool`, `datetime`,
  `date`, `time`, `timedelta`) instead of coercing everything to `str`. An
  empty cell is `None` (previously `""`).
- `read_sheet` return type is now `list[list[object]]`; trailing `None` (not
  `""`) is trimmed; `pad=True` fills with `None`.
- `Table` still coerces **row labels, column headers, and the corner** to
  `str` — you address rows and columns by name. The corner is now validated as
  `str` on construct and on `t.corner = ...` (was unconstrained).
- `grid` helpers fill gaps with `None` instead of `""` (`pad`, `transpose`,
  `get_column` of a short row, padding a short row in `set_value`/`set_column`).

### Added
- `CellTypeError(PyhandlexlError, TypeError)` — raised when a value is not a
  type Excel can store.
- `check_cell_value(value)` — validate a single value against the allowed set.
  `write_sheet` / `append_rows` / `Table.write` run it on every cell;
  timezone-aware datetimes and `Decimal` are rejected.

## [0.3.0] — 2026-09-07

### Added
- `create_workbook(path, *, sheet="Sheet")` — the only way to create a new
  `.xlsx` file.
- `delete_workbook(path)` — delete a workbook file, retrying while it is locked
  (`FileLockedError` if it stays open); refuses non-workbook paths.
- `pyhandlexl.grid` — editing helpers for the `list[list[str]]` that
  `read_sheet` returns: `set_value`, `get_row`/`get_column`, `set_row`/
  `set_column`, `insert_row`/`insert_column`, `append_row`/`append_column`,
  `delete_row`/`delete_column`, `transpose`, `pad`. Each returns a new grid;
  rows and columns are 1-based, matching `Table`.

### Changed
- **Files are never created implicitly.** `write_sheet`, `append_rows`,
  `create_sheet`, and `Table.write` now raise `FileNotFoundError` if the file
  does not exist — call `create_workbook` first.

## [0.2.3] — 2026-09-05

### Added
- `Table.add_row`, `add_column`, `rename_row`, and `rename_column` now refuse to
  create a duplicate row label or column header (`ValueError`). Constructing a
  `Table` or reading one from a file still allows duplicates.

### Changed
- Docs and package summary reframed around the `Table` API — reading and
  writing Excel worksheets as organised, labelled tables — rather than raw
  cell values.
- CI actions bumped to their Node 24 releases.

## [0.2.1] — 2026-09-05

### Added
- Published to PyPI: `pip install pyhandlexl`.

### Changed
- README rewritten; install instructions now point at PyPI.

## [0.2.0] — 2026-09-05

### Changed
- **Row labels and column headers must now be `str`.** Constructing a `Table`,
  or calling `add_row`, `add_column`, `rename_row`, or `rename_column`, with a
  non-string label/header raises `TypeError`.
- `Table.row()` / `Table.column()` renamed to `Table.read_row()` /
  `Table.read_column()`, to pair with `set_row()` / `set_column()`.
- `Table.at()` renamed and merged into `Table.read_cell()`; `Table.set()`
  renamed and merged into `Table.set_cell()`. Both now take a single
  `row=`/`column=` pair that is **either both ints** (Excel position; row 1 is
  the header row, column 1 is the label column) **or both strings** (row
  label / column header) — mixing types raises `TypeError`. A string no
  longer accepts a column letter like `"B"` for a position; use a number or
  `ref="B2"` instead.
  - `read_cell` by position can reach any cell (header, label, corner, data);
    by label it always reads data.
  - `set_cell` only ever writes **data** — addressing a header, row label, or
    the corner by position now raises `ValueError` instead of silently
    changing table structure.
- `Table.data` is now a `TableData` snapshot exposing `.rows`, `.columns`,
  `.row_labels`, `.column_headers`, and `.corner`, instead of a plain
  `list[list[str]]`.

### Removed
- `Table.at()` — use `Table.read_cell(row=..., column=...)`.
- `Table.set()` — use `Table.set_cell(row=..., column=..., value=...)`.
- `Table.row_labels` / `Table.column_headers` top-level properties — use
  `Table.data.row_labels` / `Table.data.column_headers`.
- `Table.__getitem__` (`t["Alice"]`) — use `t.read_row("Alice")`.
- `Table.__iter__` (`for row in t`) — use `for row in t.data.rows`, or
  `for column in t.data.columns`.
- `Table.__len__` (`len(t)`) — use `len(t.data.rows)`.
- `Table.__contains__` (`x in t`) — use `x in t.data.row_labels`.

## [0.1.0] — 2026-09-04

### Added
- `Table` — the labelled-table API. Row 1 is column headers, column A is row
  labels, data starts at B2. Read with `Table.read`; access by label (`at`,
  `row`, `column`) or Excel coordinate (`cell`); edit in place (`set`,
  `set_cell`, `set_row`, `set_column`, `add_row`, `add_column`, `drop_row`,
  `drop_column`, `rename_row`, `rename_column`, settable `corner`); `write` back.
- `read_sheet` / `write_sheet` / `append_rows` — raw full-grid access, with
  `orientation="columns"` and `pad` options.
- Sheet management: `list_sheets`, `sheet_exists`, `create_sheet`,
  `delete_sheet`, `rename_sheet`.
- Validation helpers: `check_sheet_name`, `check_dimensions`, `is_valid_xlsx`.
- Safe writes — every write goes to a temporary file, is verified as a readable
  `.xlsx`, then atomically replaces the original; retries while the file is
  locked before raising `FileLockedError`.
- Exception hierarchy rooted at `PyhandlexlError`.
- Continuous integration: lint and a test matrix on Python 3.10–3.13.

[Unreleased]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.0...HEAD
[0.9.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.8.0...v0.9.0
[0.8.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.7.1...v0.8.0
[0.7.1]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.7.0...v0.7.1
[0.7.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.6.0...v0.7.0
[0.6.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.2.3...v0.3.0
[0.2.3]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.2.1...v0.2.3
[0.2.1]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/LewyAmendi/pyhandlexl/releases/tag/v0.1.0
