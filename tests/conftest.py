"""Shared fixtures."""

from __future__ import annotations

import pytest

from pyhandlexl import create_file


@pytest.fixture
def book(tmp_path):
    """Path to a freshly created empty .xlsx (one worksheet, "Sheet")."""
    path = tmp_path / "book.xlsx"
    create_file(path)
    return path
