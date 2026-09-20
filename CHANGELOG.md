# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Groundwork for 1.0: an audit of the public API, and a second round of speed-ups.

### Changed (these can break existing code)
- **`nan` is now a blank cell, not an error.** `nan`, `pd.NaT`, `pd.NA` and
  `np.datetime64('NaT')` were refused (`CellTypeError`: Excel can't store them); they are
  now stored as an empty cell and read back as `None`, which is what analysis data needs.
  The infinities are still refused.
- **Formulas are no longer supported — by design, for now.** openpyxl turns any string that
  starts with `=` into a formula, and pyhandlexl used to leave it at that ("untested, may be
  removed"). That was wrong in three ways: a row label, header, corner, table name or CSV
  field like `=== NOTES ===` became a formula Excel would calculate and show as `#NAME?`
  (or reject the file over, if it wasn't a valid one); `import_csv_to_xl`, documented as
  "every field becomes a plain text cell", let a CSV plant a live `=HYPERLINK(…)`; and
  a formula written in a table went stale, silently, when rows or columns moved. So now a
  string is **always text, everywhere** — `"=1+1"` is stored as the text `=1+1`, with
  Excel's quote prefix so it stays text if someone edits the cell — and it reads back
  unchanged. Use openpyxl directly for formulas.
- **Reading a workbook that holds real formulas warns.** `read_sheet`, `Table.read` and
  `export_xl_to_csv` give a formula's *text* and raise the new **`FormulaWarning`**
  (naming the cells and pointing at your call), because writing that data back stores the
  text in place of the formula. A formula on another sheet, or outside the range you
  read and write, is untouched. Before, such a table round-tripped with its formulas
  intact.
- **One name for a row's label and a column's header on every `Table` method.**
  `read_row(row_label)`, `set_row(row_label, …)`, `read_column(column_header)` and
  `set_column(column_header, …)` now take `label` / `header`, like `add_row`, `drop_row`,
  `insert_row`, `add_column`, `drop_column` and `insert_column` always did. Positional
  calls — which is how every documented example calls them — are unaffected.
- **A missing worksheet says which.** `SheetNotFoundError` now reads
  `no worksheet named 'Data'` (it was the bare name), and it and `TableNotFoundError`
  print their message as it is instead of wrapped in the quotes `KeyError` adds. They
  are still `KeyError`s.
- **`repr(t.data)` no longer prints a whole big table.** A snapshot of more than 100
  cells shows its shape (`TableData(4000 rows x 8 columns, corner='')`); a small one shows
  its contents as before. (`rows` and `columns` hold the same cells twice, and the full
  repr of a 2,000-row table was 60 KB.)

### Added
- **Use your data with pandas and numpy.** Optional extras (`pyhandlexl[pandas]`,
  `pyhandlexl[numpy]`); the library still imports neither, and a test checks that.
  - `Table.to_dataframe()` — row labels as the index (the corner is its name), headers as the
    columns, blanks as `NaN`/`NaT`; a column restricted to a `ColumnType` is settled to the
    matching dtype (`datetime64`, `timedelta64`, `bool` or nullable `boolean`).
  - `Table.from_dataframe(df, *, name, style=None, column_types=None,
    infer_column_types=False)` — the inverse. The index and the columns must be strings (a
    table's labels are text): a non-string one is refused with a message that says how to
    fix it, and so is a `MultiIndex`. `infer_column_types=True` restricts each column to the
    type its values have.
  - `Table.to_numpy(dtype=None)` — the data as a 2-D array, with the natural dtype
    (`int64`, `float64`, `bool`, `datetime64[us]`, `timedelta64[us]`, else `object`).
  - `grid.to_dataframe(grid, header=True)` and `grid.from_dataframe(df, header=True,
    index=False)` for plain sheets.
- **numpy and pandas values are accepted wherever a value is written** (`set_cell`, `add_row`,
  `write_sheet`, `append_rows`, `check_cell_value`, …) and stored as the plain value they
  hold — `np.int64` as `int`, `np.bool_` as `bool`, `np.datetime64`/`pd.Timestamp` as
  `datetime`, `np.timedelta64`/`pd.Timedelta` as `timedelta`. Before, all of them but
  `np.float64` were refused with `CellTypeError`.
- `Table.from_dict(..., style=)`, like the constructor's.
- A **`fast` extra**: `pip install "pyhandlexl[fast]"` installs lxml, which openpyxl
  uses to write files about a quarter faster.
- `tests/test_public_api.py` pins the public API's exact surface — every exported name,
  method, property, field, exception base and signature — so from here on it can only
  change on purpose.

### Performance
- **A write repaints only what changed.** With the style and shape unchanged nothing is
  repainted (only values are rewritten); adding or dropping rows or columns repaints just
  the old and the new last row and column, the only cells whose border changes. Editing a
  cell of a 4,000 × 8 table now takes about 16% less time, and painting itself dropped
  from ~0.4 s to under 10 ms. Openpyxl's own load and save are now about two thirds of a
  write; the rest is ours. The result is cell-for-cell identical to a full repaint (a test
  replays random sequences of edits — rows and columns added, dropped and inserted, values
  of every type, style changes, a neighbouring table being shifted — both ways and
  compares every cell).
- **The well-formed-XML check on every save is four times faster** (a bare expat parser
  instead of ElementTree building a tree nobody reads), with the same verdicts, including
  for an undeclared namespace prefix.

### Fixed
- **`import_csv_to_xl`'s schema-rebuild warning pointed into the library,** at a line of
  `core.py`, not at your call (it does its work through `write_sheet`, one frame deeper
  than the depth the warning's location was computed for) — although the README says
  the warning names the line of your code. It now does, and a test covers it with the
  other operations.
- **A save that failed part-way could report the wrong error.** On Windows the half-written
  temporary file can still be open, so deleting it raised `PermissionError` and hid what
  actually went wrong (with lxml installed, a bad string made openpyxl raise
  `UnicodeEncodeError` mid-save and this got in the way). Cleanup is now best-effort.
- A cell that keeps its style across a write could keep the number format its previous
  value gave it — a number written over a date would have read back as a date. Number
  formats are reset before each value is written. (Not reachable before, because every
  write cleared every cell first.)

### Documented
- **A write drops the calculated results of formulas** (openpyxl can't keep them); the
  formulas stay and Excel recalculates on opening. The README said formulas are
  "preserved" without saying so.
- A **Limitations** section: not a database (no locks, no transactions), every write
  re-saves the whole workbook (with measured times), held in memory, Excel's data model,
  what is and isn't checked against real Excel.
- `check_sheet_name`, `check_dimensions` and `is_valid_xlsx` — public since 0.9.0, never
  mentioned in the README.
- What a write repaints, and that formatting you add by hand inside a table survives one.

## [0.9.6] — 2026-09-20

### Changed
- **Creating and rewriting tables is 6–8× faster.** Styling a table used to
  build fresh font/fill/border objects for every cell, which openpyxl then hashes to
  find each one's slot in the workbook's style table — over 90% of the time to create
  or write a table, and the reason a large table took seconds. A table has only a few
  dozen distinct looks, so each is now painted once and its style record is copied
  onto the other cells that share it. The result is identical cell for cell (a test
  compares it against the old one-cell-at-a-time painter across styles, shapes and
  positions), and the saved workbook's style table is now a few dozen entries, not
  thousands. A date, time or duration cell keeps the number format its value gave it.
  On a 4,000-row × 8-column table (one Windows machine): `create` 9.0 s → 1.2 s, and
  reading it back, adding a row and `write`-ing 11.7 s → 1.8 s.
- **`FileReadOnlyError`** (new, exported): writing to a read-only file now raises this
  instead of a plain `FileLockedError`, because the two call for different responses
  (retrying can never help a read-only file). It *is* a `FileLockedError`, so code that
  catches that keeps working unchanged.
- **`MergeReport` is now `pyhandlexl.MergeReport`** as far as `repr`, `help()`, pickle
  and the docs are concerned, instead of showing a private module.
- **openpyxl `>=3.1.3,<4`.** 3.1.0–3.1.2 leak a file handle on Windows, which surfaced
  as a spurious "file is locked" error after a save. The lower bound is now tested
  in CI (on Linux and Windows) alongside the newest release, and 4.x is excluded until
  it has been tried.
- **Status is now Beta** (PyPI classifier), and Python 3.14 is supported and tested,
  as is macOS.
- **The package is checked with `mypy --strict` in CI** (it already shipped `py.typed`).
  A few internals were retyped to satisfy it; no public signature changed.

### Documented
- **Formulas are not a supported feature.** A string starting with `=` is still
  stored as a formula (an openpyxl behaviour), but that is untested and **may change
  or be removed in a future release**; the README now says so where it describes it.
- **A CSV stays all text** — every value is a `str`, nothing is inferred. The README
  no longer says this "may change in a future release"; it won't.

## [0.9.5] — 2026-09-20

### Changed
- **`Table.read` says which cell is at fault** when a header or row label is blank
  (`table 'T' on sheet 'D' can't be read: the column header in cell C2 is blank …`),
  instead of an anonymous `column headers must not be empty`.
- **Every save now checks that the file it wrote is well-formed XML,** not just that
  it opens as a workbook, before it replaces the original (the old check read
  worksheets lazily, so a poisoned sheet passed it).
- **The schema-rebuild warning now points at your code.** `SchemaRebuiltWarning`
  used to be attributed to a line inside pyhandlexl (`table.py`, `core.py`), so
  you couldn't tell which of your calls had triggered it — and Python's default
  once-per-location filtering could hide a repeat. It now reports the line that
  called the pyhandlexl function.
- **Every operation that rebuilds the schema now saves the rebuild.**
  `sheet_kind`, `write_sheet` (to an existing sheet) and `append_rows` used to
  warn but leave the schema sheet missing, so they rescanned and warned again on
  every call. They now save it, like `Table.read` and `list_tables` already did.
- The README's schema-rebuild section is now two explicit lists — what a rebuild
  restores, and what it does not (and cannot) — including which operations trigger
  it and that a table whose marker is gone is not recovered.

### Fixed
- **One bad string could make an entire workbook unreadable.** A string containing a
  lone surrogate (`"\ud800"`) or `"\ufffe"`/`"\uffff"` was written without complaint;
  the result could not be loaded again — not even the *other* sheets — while
  `is_valid_xlsx` still said yes and the good file had already been replaced. A control
  character raised a raw openpyxl `IllegalCharacterError` instead of a pyhandlexl
  error. All of these now raise `CellTypeError` up front — in values, table names,
  labels, headers and the corner alike.
- **Values Excel silently changed or lost are refused:** `nan`/`inf` (became a blank
  cell), strings over 32,767 characters (silently truncated — and a 32,768-character
  *table name* left the table unfindable), dates before 1900-01-01 or after
  9999-12-31 (`datetime.max` came back as the text `#VALUE!`), and ints beyond float
  range. `CellTypeError`.
- **Line breaks were stored differently on different platforms:** `"a\r\nb"` came
  back as `"a\n\nb"` on Windows (a blank line gained) and `"a\nb"` on Linux. Data
  cells now store `\n`, everywhere; a carriage return in a table's name, label,
  header or corner (found by exact text) is refused.
- **Macro-enabled workbooks lost their macros.** Any write to an `.xlsm`/`.xltm`
  silently stripped `vbaProject.bin` and left a file Excel would not open. The macros
  are now kept, and openpyxl's stray `Exception ignored … ZipFile.__del__` message
  when handling one is silenced.
- **`create_workbook` accepted any file name** — `data.csv`, `book.xlsm`, `book.xls`,
  no extension — creating `.xlsx` content Excel refuses to open under that name (and
  which `write_sheet` then mistook for a CSV). It now makes only `.xlsx` files, and
  names a missing directory instead of a temporary file.
- **A table with more than ~5,400 columns locked the whole workbook.** Its column
  types are stored in one cell, which held at most 32,767 characters; the overflow was
  cut off and every later operation died with `JSONDecodeError`. Types are now stored
  compactly when they wouldn't fit; tables up to Excel's 16,383 data columns work.
- **A rebuilt schema mistook data for tables.** A row label, header, corner, or data
  value reading `TABLE NAME` inside a table became a phantom table (and cut the real
  table's width short, so it stopped reading). Cells inside an already-found table
  are no longer treated as markers.
- **A bare string was silently split into characters** by `add_row`, `insert_row`,
  `set_row`, `set_column`, `add_column`, `insert_column`, `write_sheet`,
  `append_rows`, and the `pyhandlexl.grid` functions. It now raises `TypeError`, as
  the `Table` constructor already did.
- **On Linux and macOS a save overwrote read-only files, reset their permissions, and
  replaced symlinks with regular files.** A save writes a new file and swaps it in,
  which on POSIX needs write access to the *directory* only — so a `chmod 444` file
  was silently replaced (Windows refused), its mode went back to the default, and a
  symlinked workbook became an ordinary file while its real target stayed stale. A
  read-only file (`.xlsx` or `.csv`) now raises `FileLockedError` on every platform
  before anything is written, permissions are kept, and a symlink is followed so its
  target is the file that is updated.
- **A CSV containing a NUL byte raised a bare `_csv.Error` on Python 3.10** (3.11+
  reads it). It is now a `ValueError` naming the line.
- **Reading a CSV with a field over 131,072 characters raised `_csv.Error`,** despite
  the README saying a CSV has no size limit.
- **Sheet names with control characters** (or a leading/trailing apostrophe, which
  Excel forbids) were accepted and then failed with a confusing `InvalidFileError`;
  they raise `SheetNameError` now.
- **A sheet name that collides with the reserved schema sheet's is refused.**
  `_PYHANDLEXL_TABLES` (any capitalisation, or with surrounding spaces) used to be
  accepted by `create_sheet`, `rename_sheet`, `write_sheet`, `append_rows` and
  `Table.create`. Excel ignores case and openpyxl silently renames a would-be
  duplicate, so the call reported success but produced `…_TABLES1`, and a lookalike
  created before the schema existed pushed the real schema sheet to
  `_pyhandlexl_tables1` where it was never found again. All of these now raise
  `SheetKindError`, as does `clear_all_sheet_data`/`delete_sheet`/`sheet_kind`, and
  `create_sheet` refuses the exact reserved name even when no schema exists yet.
  Saving the schema into a hand-edited file that already holds such a lookalike now
  raises `SheetKindError` instead of quietly misnaming it.
- **A sheet name that differs from an existing one only by capitalisation is
  refused, not silently renamed.** `create_sheet("DATA")` with `Data` present used
  to succeed and create `DATA1`; `write_sheet(sheet="data")` wrote to a brand-new
  `data1` sheet instead of `Data`; `rename_sheet` did the same. These now raise
  `ValueError` naming the clashing sheet. Renaming a sheet's own capitalisation
  (`log` → `LOG`) now works — openpyxl used to turn it into `LOG1`.

## [0.9.4] — 2026-09-19

### Added
- **`Table.write()` merges concurrent changes instead of overwriting them.**
  A `Table` remembers what it last saw on disk; if another writer changed the
  table since, `write()` performs a three-way merge keyed by row label and
  column header. Non-overlapping changes combine silently (rows and columns
  each side added, edits to different cells, deletions the other side left
  alone, renames, corner/style/column-type changes only one side made); where
  both sides changed the same thing, **yours wins** and a `MergeConflictWarning`
  lists what it overwrote. Deleting a row or column the other writer edited, or
  editing one they deleted, is reported the same way. Rows and columns you both
  appended keep their order (theirs, then yours). Afterwards the object holds
  the merged table.
  - `t.last_merge` (new) — a `MergeReport` (`rows_added`, `rows_removed`,
    `columns_added`, `columns_removed`, `cells_updated`, `conflicts`) of what the
    last `write()` merged in, or `None` if nothing had changed on disk.
  - `MergeConflictWarning` and `MergeReport` (new, exported). A
    `MergeConflictWarning` is a `Warning`, not a `PyhandlexlError`; the write still
    succeeds.
  - Column types are checked *after* merging, so a restriction another writer
    added can reject a value you set (`ColumnTypeError`); nothing is written and
    the `Table` object is left exactly as it was. A `Table` never read or created
    has nothing to merge against and overwrites as before; a table with duplicate
    labels falls back to yours replacing the disk version, with a warning.
  - **This is not a lock.** The merge covers the time between `read` and
    `write`; it does not stop two writes landing at the very same instant, nor
    writes to different tables in one workbook overwriting each other (each write
    re-saves the whole workbook). No file locking is used, since OS-level locking
    behaves differently across platforms and filesystems.

### Changed
- **Schema rebuild now infers column types instead of resetting them all
  to `ColumnType.ANY`.** When the reserved `_pyhandlexl_tables` sheet is
  deleted and rebuilt from markers, each data column is scanned: if every
  non-blank value belongs to one Excel-native type, the column comes back
  restricted to it (`TEXT`, `NUMBER`, `BOOLEAN`, `DATE`, `TIME`, or
  `DURATION`); an empty or mixed column comes back `ANY`. The inferred
  types always admit the data already in the table, so writing a rebuilt
  table straight back never raises `ColumnTypeError`.
  - This is inference, not recovery — nothing in a cell records a
    restriction — so it can go either way: an empty or mixed column that
    was restricted comes back `ANY` (lost), and a column deliberately left
    unrestricted, but holding a single type, comes back restricted to it
    (invented). Check `t.column_types` after a rebuild.
  - `SchemaRebuiltWarning`'s message and the schema sheet's warning comment
    now say this instead of "cannot be recovered at all". Creation and
    modification times are still unrecoverable (there's no data to infer
    them from) and still come back `None`.

### Fixed
- **Workbooks written by older versions no longer fail to load.** A reserved
  schema sheet from ≤ 0.9.2 (7 or 8 columns, before column types and
  creation/modification times existed) raised `ValueError`. It now loads with
  defaults (default style, unrestricted columns, `None` dates) and is upgraded to
  the current layout by the next write.

## [0.9.3] — 2026-09-17

### Added
- **A table now tracks its own creation/modification times, size, and
  sheet — accessible via `t.info` and the new standalone `table_info`.**
  - `t.info` (new property) — a `TableInfo` snapshot: `created_at`,
    `modified_at` (both UTC `datetime`, set by `create()`/`write()`
    respectively), `n_rows`, `n_cols`, `sheet`, plus `style` and
    `column_types` cross-referenced from the existing properties of the
    same name. `sheet`/`created_at`/`modified_at` are `None` until the
    table has actually been placed with `create()` (or loaded with
    `read()`).
  - `table_info(path, name)` (new, exported) — the same `TableInfo`
    without constructing a `Table` or reading its row data; useful for
    cheaply checking size/dates/sheet across many tables.
  - `TableInfo` (new, exported) — the frozen dataclass both return.
  - `modified_at` reflects only a table's *own* `create()`/`write()` —
    being shifted right to make room for a growing neighbor doesn't count,
    since the table's own data didn't change.
  - Persisted per table in the reserved schema sheet, alongside style and
    column types — and subject to the same rebuild limitation: if that
    sheet is deleted and reconstructed from markers, there's nothing in a
    cell recording either time, so a rebuilt table reports both as `None`
    rather than a guessed value. Size and sheet still come back exact.

## [0.9.2] — 2026-09-17

### Added
- **A column can now be restricted to a single Excel-native data type.**
  Every column defaults to `ColumnType.ANY` (anything `check_cell_value`
  allows) unless declared otherwise; a value that doesn't match its
  column's type raises `ColumnTypeError` on the next `create()`/`write()`
  — the same schedule as `CellTypeError`, not when values are set. A blank
  cell (`None`) is always allowed regardless of a column's type.
  - `ColumnType` (new, exported) — an enum of `ANY`, `TEXT`, `NUMBER`,
    `BOOLEAN`, `DATE`, `TIME`, `DURATION`. Follows Excel's own type model,
    not Python's: `NUMBER` covers both `int` and `float`, and `DATE`
    covers both `datetime.date` and `datetime.datetime`.
  - `Table(..., column_types={"header": ColumnType.NUMBER, ...})` — declare
    at construction; `t.column_types` reads the current mapping (every
    column, defaulting to `ANY`); `t.set_column_type(header, column_type)`
    changes one afterward. `Table.from_dict(..., column_types=...)` too.
  - Persisted per table in the reserved `_pyhandlexl_tables` schema sheet.
    Renaming, inserting, or dropping a column moves or drops its
    restriction along with it; a newly inserted column always starts as
    `ANY`. Included in `Table` equality (unlike `style`, which isn't).
  - `ColumnTypeError` (new, exported) — a `PyhandlexlError` and `TypeError`.
  - Like a table's style, a column's type restriction can't be recovered if
    the schema sheet is deleted and rebuilt from markers — a rebuilt table
    always reports `ColumnType.ANY` for every column.

- **The reserved `_pyhandlexl_tables` schema sheet is now visibly marked**
  — a red sheet tab and a warning comment on its first cell, written every
  time the schema is saved (including right after a rebuild) — so it reads
  as "don't touch this" to anyone opening the workbook by hand, not just to
  someone who's read the docs. `delete_sheet` and `rename_sheet` now back
  that up: deleting the schema sheet directly, renaming it away, or
  renaming some other sheet onto its reserved name are all refused
  (`SheetKindError`) rather than silently allowed — it's only ever meant
  to disappear (or move) as a side effect of the tables that live on it,
  not as a direct target.

- **A worksheet now holds either named tables or plain grid data, never
  both.** Whichever writes to a sheet first claims it: `write_sheet`/
  `append_rows` claim it as `"grid"`; `Table.create` claims it as `"table"`.
  Writing the other kind to an already-claimed sheet now raises
  `SheetKindError` instead of silently risking corruption (a grid write
  landing across a table's marker, or a table placed over unrelated grid
  content). The reserved `_pyhandlexl_tables` schema sheet can't be
  targeted directly by any of them either.
  - `sheet_kind(path, sheet)` (new, exported) — `"table"`, `"grid"`, or
    `"empty"` for a given sheet.
  - `clear_all_sheet_data(path, sheet)` (new, exported) — wipes a sheet's
    cells, styles, and any tables tracked on it back to `"empty"`, the one
    way to reclaim it as the other kind.
  - `SheetKindError` (new, exported) — a `PyhandlexlError` and `ValueError`.

### Fixed
- **`delete_sheet` and `rename_sheet` now keep the schema in sync.**
  Deleting a sheet that held tables used to leave their schema entries
  stale (pointing at a sheet that no longer exists); it now forgets them.
  Renaming a table-holding sheet used to leave entries pointing at the old
  name; it now updates them, so the tables stay readable under their same
  names.

## [0.9.1] — 2026-09-15

### Changed
- **`TableStyle.DEFAULT`'s header/label look**: black text (was white) on an
  olive green fill (was blue) — `header_font_color="000000"`,
  `header_fill="76933C"` — for better contrast. `TableStyle.MINIMAL` no
  longer needs to override the font color to get there itself.
- **`list_sheets()` no longer includes the reserved `_pyhandlexl_tables`
  schema sheet** — it isn't a sheet you created or can write to, so it's
  filtered out rather than shown like any other sheet.
- **`Table.corner` is no longer a property.** Set it with
  `t.set_corner(value)` (still `str`-only, still raises `TypeError`
  otherwise); read it through `t.data.corner`, same as every other field on
  the read-only snapshot.
- **Column headers and row labels can no longer be `""`.** Constructing a
  `Table`, or calling `add_row`/`add_column`/`insert_row`/`insert_column`/
  `rename_row`/`rename_column` with an empty one, now raises `ValueError`.
  The corner is unaffected — it may still be `""`. The automatic schema
  rebuild (see 0.9.0 below) still tolerates a blank header/label already
  sitting in a hand-edited or pre-existing sheet; it just can't be created
  through the API anymore.
- **`read_sheet` gained an `orientation` parameter**, matching
  `write_sheet`: `"rows"` (default, unchanged) returns each worksheet row as
  an inner list; `"columns"` returns each worksheet column as an inner list
  instead — the transpose of `"rows"`, with the same trailing-empty
  trimming/`pad` rules applied down each column.

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

[Unreleased]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.6...HEAD
[0.9.6]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.5...v0.9.6
[0.9.5]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.9.0...v0.9.1
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
