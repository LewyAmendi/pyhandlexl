"""pyhandlexl — read and write raw cell values in Excel .xlsx files."""

from pyhandlexl.core import append_rows, read_sheet, write_sheet
from pyhandlexl.errors import (
    DimensionError,
    FileLockedError,
    InvalidFileError,
    PyhandlexlError,
    SheetNameError,
    SheetNotFoundError,
)
from pyhandlexl.validate import check_dimensions, check_sheet_name, is_valid_xlsx

__version__ = "0.1.0"

__all__ = [
    "DimensionError",
    "FileLockedError",
    "InvalidFileError",
    "PyhandlexlError",
    "SheetNameError",
    "SheetNotFoundError",
    "append_rows",
    "check_dimensions",
    "check_sheet_name",
    "is_valid_xlsx",
    "read_sheet",
    "write_sheet",
]
