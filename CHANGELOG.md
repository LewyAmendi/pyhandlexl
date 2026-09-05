# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `Table.add_row`, `add_column`, `rename_row`, and `rename_column` now refuse to
  create a duplicate row label or column header (`ValueError`). Constructing a
  `Table` or reading one from a file still allows duplicates.

### Changed
- Docs and package summary reframed around the `Table` API — reading and
  writing Excel worksheets as organised, labelled tables — rather than raw
  cell values.

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

[Unreleased]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.2.1...HEAD
[0.2.1]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/LewyAmendi/pyhandlexl/releases/tag/v0.1.0
