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
