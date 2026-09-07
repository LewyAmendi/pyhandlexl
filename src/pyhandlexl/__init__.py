"""pyhandlexl — read and write Excel .xlsx worksheets as organised, labelled tables."""

from pyhandlexl import grid
from pyhandlexl.core import (
    append_rows,
    create_sheet,
    create_workbook,
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
from pyhandlexl.table import Table, TableData
from pyhandlexl.validate import check_dimensions, check_sheet_name, is_valid_xlsx

__version__ = "0.3.0"

__all__ = [
    "DimensionError",
    "FileLockedError",
    "InvalidFileError",
    "PyhandlexlError",
    "SheetNameError",
    "SheetNotFoundError",
    "Table",
    "TableData",
    "append_rows",
    "check_dimensions",
    "check_sheet_name",
    "create_sheet",
    "create_workbook",
    "delete_sheet",
    "grid",
    "is_valid_xlsx",
    "list_sheets",
    "read_sheet",
    "rename_sheet",
    "sheet_exists",
    "write_sheet",
]
