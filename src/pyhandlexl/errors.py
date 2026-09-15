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


class SheetNotFoundError(PyhandlexlError, KeyError):
    """No worksheet with the requested name exists in the workbook."""


class CellTypeError(PyhandlexlError, TypeError):
    """A cell value is not a type Excel can store."""


class TableNotFoundError(PyhandlexlError, KeyError):
    """No named table with this name exists, or its marker cannot be found."""


class TableExistsError(PyhandlexlError, ValueError):
    """A named table with this name already exists in the workbook."""


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
