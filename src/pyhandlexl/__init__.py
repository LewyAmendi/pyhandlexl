"""pyhandlexl — read and write raw cell values in Excel .xlsx files."""

from pyhandlexl.core import (
    append_rows,
    create_sheet,
    delete_sheet,
    list_sheets,
    read_sheet,
    rename_sheet,
    sheet_exists,
    write_sheet,
)
from pyhandlexl.errors import (
    DimensionError,
    FileLockedError,
    InvalidFileError,
    PyhandlexlError,
    SheetNameError,
    SheetNotFoundError,
)
from pyhandlexl.table import Table
from pyhandlexl.validate import check_dimensions, check_sheet_name, is_valid_xlsx

__version__ = "0.1.0"

__all__ = [
    "DimensionError",
    "FileLockedError",
    "InvalidFileError",
    "PyhandlexlError",
    "SheetNameError",
    "SheetNotFoundError",
    "Table",
    "append_rows",
    "check_dimensions",
    "check_sheet_name",
    "create_sheet",
    "delete_sheet",
    "is_valid_xlsx",
    "list_sheets",
    "read_sheet",
    "rename_sheet",
    "sheet_exists",
    "write_sheet",
]
