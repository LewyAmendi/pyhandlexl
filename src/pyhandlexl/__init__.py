"""pyhandlexl — read and write Excel .xlsx worksheets as organised, labelled tables."""

from pyhandlexl import grid
from pyhandlexl.core import (
    append_rows,
    create_csv,
    create_sheet,
    create_workbook,
    delete_sheet,
    delete_table,
    delete_workbook,
    export_xl_to_csv,
    import_csv_to_xl,
    list_sheets,
    list_tables,
    read_sheet,
    rename_sheet,
    sheet_exists,
    write_sheet,
)
from pyhandlexl.errors import (
    CellTypeError,
    DimensionError,
    FileLockedError,
    InvalidFileError,
    PyhandlexlError,
    SheetNameError,
    SheetNotFoundError,
    TableExistsError,
    TableNotFoundError,
)
from pyhandlexl.table import Table, TableData
from pyhandlexl.validate import (
    check_cell_value,
    check_dimensions,
    check_sheet_name,
    is_valid_xlsx,
)

__version__ = "0.8.0"

__all__ = [
    "CellTypeError",
    "DimensionError",
    "FileLockedError",
    "InvalidFileError",
    "PyhandlexlError",
    "SheetNameError",
    "SheetNotFoundError",
    "Table",
    "TableData",
    "TableExistsError",
    "TableNotFoundError",
    "append_rows",
    "check_cell_value",
    "check_dimensions",
    "check_sheet_name",
    "create_csv",
    "create_sheet",
    "create_workbook",
    "delete_sheet",
    "delete_table",
    "delete_workbook",
    "export_xl_to_csv",
    "grid",
    "import_csv_to_xl",
    "is_valid_xlsx",
    "list_sheets",
    "list_tables",
    "read_sheet",
    "rename_sheet",
    "sheet_exists",
    "write_sheet",
]
