"""Exception types raised by pyhandlexl."""


class PyhandlexlError(Exception):
    """Base class for every error raised by pyhandlexl.

    Catch this to handle any failure from the library.
    """


class SheetNameError(PyhandlexlError, ValueError):
    """A worksheet name is empty, too long, or contains illegal characters."""


class DimensionError(PyhandlexlError, ValueError):
    """The data has more rows or columns than the .xlsx format allows."""


class InvalidFileError(PyhandlexlError):
    """The file is missing, not a zip, or not a readable .xlsx workbook."""


class FileLockedError(PyhandlexlError, OSError):
    """The file stayed locked (e.g. open in Excel) after every retry."""


class FileReadOnlyError(FileLockedError):
    """The file is read-only, so nothing was written.

    A kind of :class:`FileLockedError` — code that already catches that keeps working —
    but retrying can't help: the file has to be made writable first.
    """


class SheetNotFoundError(PyhandlexlError, KeyError):
    """No worksheet with the requested name exists in the workbook."""

    def __str__(self) -> str:  # KeyError would show repr(message), quotes and all
        return Exception.__str__(self)


class CellTypeError(PyhandlexlError, TypeError):
    """A cell value is not a type Excel can store."""


class ColumnTypeError(PyhandlexlError, TypeError):
    """A data value doesn't match its column's declared :class:`~pyhandlexl.column_type.ColumnType`.

    Raised by ``Table.write``/``Table.create``, the same time as
    ``CellTypeError`` — a value can pass ``CellTypeError`` (it's a type
    Excel can store at all) and still fail this (it isn't the type this
    particular column was restricted to). ``None`` never triggers this,
    regardless of the column's restriction.
    """


class TableNotFoundError(PyhandlexlError, KeyError):
    """No named table with this name exists, or its marker cannot be found."""

    def __str__(self) -> str:  # KeyError would show repr(message), quotes and all
        return Exception.__str__(self)


class TableExistsError(PyhandlexlError, ValueError):
    """A named table with this name already exists in the workbook."""


class SheetKindError(PyhandlexlError, ValueError):
    """A worksheet already holds the other kind of data.

    A sheet holds either named tables or plain grid data, never both — the
    first successful write claims it. ``write_sheet``/``append_rows`` refuse
    a sheet that already holds a table (use ``Table.write``/``Table.create``
    instead), and ``Table.create`` refuses a sheet that already holds plain
    grid data (use ``write_sheet``/``append_rows`` instead). The reserved
    ``_pyhandlexl_tables`` schema sheet can't be targeted by any of these
    directly either way. Call ``clear_all_sheet_data`` to wipe a sheet back
    to unclaimed so it can be written as the other kind.
    """


class MergeConflictWarning(UserWarning):
    """``Table.write`` found the table changed on disk since it was read, merged
    those changes with yours, and had to overwrite some of them.

    Not a ``PyhandlexlError`` — the write still succeeded. Changes that don't
    overlap (different rows, different cells, both sides appending) merge
    silently; this fires only where both sides changed the same thing
    differently (the same cell, the same new row or column, an edit to
    something the other side deleted, and so on). Your side always wins, and
    the message lists what it overwrote.
    """


class FormulaWarning(UserWarning):
    """A workbook being read holds formulas, which pyhandlexl does not support.

    Not a ``PyhandlexlError`` — the read still succeeds. A formula cell is read as its
    text (``"=A1+B1"``), and pyhandlexl never writes a formula: writing that text back
    stores it as text, replacing the formula. So read a range that holds formulas only if
    you mean to turn them into text; use openpyxl directly to work with formulas.
    """


class SchemaRebuiltWarning(UserWarning):
    """The reserved schema sheet was missing and has been reconstructed by
    scanning the workbook for table markers.

    Not a ``PyhandlexlError`` — the operation that triggered this still
    succeeds. Geometry (position and size) is exact, reconstructed from the
    markers themselves; a table's :class:`~pyhandlexl.style.TableStyle` is
    read back from its painted cells on a best-effort basis and may not
    exactly match what was originally set, especially if formatting was
    edited by hand afterwards.
    """
