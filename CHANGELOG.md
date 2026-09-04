# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/LewyAmendi/pyhandlexl/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/LewyAmendi/pyhandlexl/releases/tag/v0.1.0
