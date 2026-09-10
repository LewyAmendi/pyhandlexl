"""Tests for the exception hierarchy."""

import pytest

from pyhandlexl.errors import (
    CellTypeError,
    DimensionError,
    FileLockedError,
    InvalidFileError,
    PyhandlexlError,
    SheetNameError,
    SheetNotFoundError,
)


def test_all_errors_derive_from_base():
    for exc in (
        SheetNameError,
        DimensionError,
        InvalidFileError,
        FileLockedError,
        SheetNotFoundError,
        CellTypeError,
    ):
        assert issubclass(exc, PyhandlexlError)


def test_cell_type_error_can_be_caught_as_type_error():
    with pytest.raises(TypeError):
        raise CellTypeError("list is not a type Excel can store")


def test_file_locked_error_can_be_caught_as_oserror():
    with pytest.raises(OSError):
        raise FileLockedError("still open in Excel")


def test_sheet_name_error_can_be_caught_as_value_error():
    with pytest.raises(ValueError):
        raise SheetNameError("illegal character in name")


def test_dimension_error_can_be_caught_as_value_error():
    with pytest.raises(ValueError):
        raise DimensionError("too many rows")


def test_sheet_not_found_error_can_be_caught_as_key_error():
    with pytest.raises(KeyError):
        raise SheetNotFoundError("Sales")


def test_invalid_file_error_is_base_only():
    assert issubclass(InvalidFileError, PyhandlexlError)
    assert not issubclass(InvalidFileError, (ValueError, OSError, KeyError))
