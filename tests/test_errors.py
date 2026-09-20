"""Tests for the exception hierarchy."""

import pytest

from pyhandlexl.errors import (
    CellTypeError,
    DimensionError,
    FileLockedError,
    FileReadOnlyError,
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
        FileReadOnlyError,
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


def test_a_read_only_error_is_a_locked_error_so_older_handlers_still_catch_it():
    assert issubclass(FileReadOnlyError, FileLockedError)
    with pytest.raises(FileLockedError):
        raise FileReadOnlyError("read-only")
    with pytest.raises(OSError):
        raise FileReadOnlyError("read-only")


def test_sheet_name_error_can_be_caught_as_value_error():
    with pytest.raises(ValueError):
        raise SheetNameError("illegal character in name")


def test_dimension_error_can_be_caught_as_value_error():
    with pytest.raises(ValueError):
        raise DimensionError("too many rows")


def test_sheet_not_found_error_can_be_caught_as_key_error():
    with pytest.raises(KeyError):
        raise SheetNotFoundError("Sales")


def test_invalid_file_error_is_a_value_error_like_every_parse_failure():
    # (json.JSONDecodeError and tomllib.TOMLDecodeError are ValueErrors too — and an unreadable CSV
    # was a ValueError before it became an InvalidFileError)
    assert issubclass(InvalidFileError, PyhandlexlError)
    assert issubclass(InvalidFileError, ValueError)
    assert not issubclass(InvalidFileError, (OSError, KeyError))
