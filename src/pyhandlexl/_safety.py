"""Resilient workbook I/O: retry-on-lock loading and atomic saving.

Internal module — not part of the public API.
"""

from __future__ import annotations

import gc
import importlib
import os
import shutil
import xml.etree.ElementTree as ET
import zlib
from collections.abc import Callable, Iterator
from contextlib import contextmanager, suppress
from pathlib import Path
from secrets import token_hex
from time import sleep
from typing import TypeVar
from xml.parsers import expat
from zipfile import BadZipFile, ZipFile, is_zipfile

from openpyxl import Workbook, load_workbook

from pyhandlexl.errors import (
    FileLockedError,
    FileReadOnlyError,
    InvalidFileError,
    PyhandlexlError,
)
from pyhandlexl.validate import is_valid_xlsx

_T = TypeVar("_T")

DEFAULT_RETRIES = 5
DEFAULT_DELAY = 0.5

# Raised when a file is briefly unavailable: open in Excel, or a competing
# writer has not finished flushing it yet. Worth retrying.
_TRANSIENT_ERRORS = (PermissionError, BadZipFile, EOFError)

# Macro-enabled files: without ``keep_vba`` openpyxl drops the macros on load, so a plain
# read-modify-save would silently strip them — and the file, still called .xlsm, would no
# longer open in Excel.
_MACRO_SUFFIXES = frozenset({".xlsm", ".xltm"})


def resolve_target(path: str | Path) -> Path:
    """The real file behind *path*.

    Saving replaces the file with a freshly written one. Left to itself that would swap a
    symbolic link for a regular file (leaving the file it pointed at untouched and stale),
    so a link is followed first and its target is what gets replaced.
    """
    return Path(os.path.realpath(path))


def ensure_writable(path: Path) -> None:
    """Refuse to replace a read-only file.

    A save works by writing a new file beside the old one and swapping it in, which on
    POSIX needs write access to the *directory* only — so without this check a file made
    read-only on purpose (``chmod 444``) would be silently overwritten on Linux and macOS,
    while Windows refuses. This makes them agree, and refuses before doing any work.
    """
    if path.exists() and not os.access(path, os.W_OK):
        raise FileReadOnlyError(f"{path} is read-only (Permission denied); nothing was written")


def keep_permissions(original: Path, replacement: Path) -> None:
    """Give *replacement* the permission bits of the file it is about to replace."""
    try:
        if original.exists():
            shutil.copymode(original, replacement)
    except OSError:  # best effort: a filesystem without modes just keeps the default
        pass


def _is_well_formed(path: Path) -> bool:
    """Whether every XML part of the .xlsx at *path* parses.

    ``is_valid_xlsx`` opens the workbook read-only, which reads worksheets lazily — so
    a sheet holding an XML-illegal character passes it and would replace a good file.
    This streams each part through a real XML parser, which is what actually catches it.
    The bare expat parser is what ElementTree itself is built on (namespaces on, as there),
    without building a tree nobody reads — about four times faster on a big sheet.
    """
    try:
        with ZipFile(path) as archive:
            for name in archive.namelist():
                if name.endswith((".xml", ".rels")):
                    with archive.open(name) as part:
                        expat.ParserCreate(None, "}").ParseFile(part)
    except (expat.ExpatError, BadZipFile, OSError):
        return False
    return True


def _retry(
    operation: Callable[[], _T],
    *,
    path: Path,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> _T:
    """Run *operation*, retrying on transient file errors.

    Raises FileLockedError if it never succeeds.
    """
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            return operation()
        except _TRANSIENT_ERRORS as error:
            last_error = error
            if attempt < retries:
                sleep(delay)
    raise FileLockedError(
        f"could not access {path} after {retries} attempt(s): {last_error}"
    ) from last_error


def safe_load(
    path: str | Path,
    *,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> Workbook:
    """Load an existing workbook, retrying while the file is locked.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file exists but is not a readable .xlsx.
        FileLockedError: the file stayed locked through every retry.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"no file at {path}")
    if not is_valid_xlsx(path):
        raise InvalidFileError(f"{path} is not a readable .xlsx workbook")
    keep_vba = path.suffix.lower() in _MACRO_SUFFIXES
    with damage_is_invalid(path):  # a damaged sheet is an invalid file, not a raw XML error
        workbook = _retry(
            lambda: load_workbook(path, keep_vba=keep_vba), path=path, retries=retries, delay=delay
        )
    if keep_vba:
        _close_vba_archive_with(workbook)
    return workbook


# Reads open a workbook read-only, so only the sheets they touch are ever parsed — a full
# load parses every sheet, which made reading one small table from a workbook of big sheets
# cost as much as reading all of them. This is always True; it exists so a test can compare
# every read against a full load.
READ_ONLY_READS = True


def _xml_errors() -> tuple[type[BaseException], ...]:
    errors: list[type[BaseException]] = [ET.ParseError, expat.ExpatError]
    with suppress(ImportError):  # openpyxl uses lxml, when it is installed, to parse
        errors.append(importlib.import_module("lxml.etree").LxmlError)
    return tuple(errors)


# What a damaged part of the file raises when it is read: a bad checksum, a truncated or
# missing part, XML that doesn't parse.
_DAMAGE: tuple[type[BaseException], ...] = (
    BadZipFile,
    EOFError,
    KeyError,
    zlib.error,
    UnicodeDecodeError,
    *_xml_errors(),
)


def safe_load_readonly(
    path: str | Path,
    *,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> Workbook:
    """Open an existing workbook for **reading only**, parsing a sheet only when it is read.

    Cheaper than :func:`safe_load` in proportion to the sheets you don't read, and it skips
    that function's pre-check, which decompresses the whole archive. The workbook is a
    stream over the file: close it when done, and don't save it. A damaged part is not
    noticed until it is read — wrap the reading in :func:`damage_is_invalid` to have that
    come out as ``InvalidFileError``.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file exists but is not a readable .xlsx.
        FileLockedError: the file stayed locked through every retry.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"no file at {path}")
    if not path.is_file() or not is_zipfile(path):  # not worth retrying: it will never open
        raise InvalidFileError(f"{path} is not a readable .xlsx workbook")
    try:
        return _retry(
            lambda: load_workbook(path, read_only=True, keep_links=False),
            path=path,
            retries=retries,
            delay=delay,
        )
    except FileLockedError:
        raise
    except Exception as error:
        raise InvalidFileError(f"{path} is not a readable .xlsx workbook") from error


def open_for_reading(path: str | Path) -> Workbook:
    """The workbook to read from: a read-only stream (see :func:`safe_load_readonly`), or a
    full load when a test has switched that off."""
    return safe_load_readonly(path) if READ_ONLY_READS else safe_load(path)


@contextmanager
def damage_is_invalid(path: str | Path) -> Iterator[None]:
    """Turn a damaged part of a file, met while reading it, into ``InvalidFileError``.

    (A pyhandlexl error raised inside, such as ``TableNotFoundError`` — itself a
    ``KeyError`` — passes through untouched.)
    """
    try:
        yield
    except PyhandlexlError:
        raise
    except _DAMAGE as error:
        raise InvalidFileError(
            f"{path} is not a readable .xlsx workbook ({type(error).__name__}: {error})"
        ) from error


def _close_vba_archive_with(workbook: Workbook) -> None:
    """Make ``workbook.close()`` also close the macro archive openpyxl keeps beside it.

    With ``keep_vba`` openpyxl holds the macros in a zip backed by a temporary in-memory
    buffer. Left to garbage collection, the buffer can be finalised before the zip, and
    the zip's ``__del__`` then prints "Exception ignored ... I/O operation on closed
    file" to stderr — once per load. Closing it ourselves, in order, avoids that.
    """
    original_close = workbook.close

    def close() -> None:
        original_close()
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None and archive.fp is not None:
            try:
                archive.close()
            except ValueError:  # its buffer is already gone: just stop it trying again
                archive.fp = None

    workbook.close = close  # type: ignore[method-assign]


def safe_delete(
    path: str | Path,
    *,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> None:
    """Delete a file, retrying while it is locked.

    Raises:
        FileNotFoundError: no file at *path*.
        FileLockedError: the file stayed locked through every retry.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"no file at {path}")
    _retry(path.unlink, path=path, retries=retries, delay=delay)


def safe_rename(
    path: str | Path,
    target: str | Path,
    *,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> None:
    """Rename a file to *target*, retrying while it is locked; never overwrites.

    Raises:
        FileNotFoundError: no file at *path*.
        FileExistsError: something else is already at *target*.
        FileLockedError: the file stayed locked through every retry.
    """
    path, target = Path(path), Path(target)
    if not path.exists():
        raise FileNotFoundError(f"no file at {path}")
    if target.exists() and not os.path.samefile(path, target):
        raise FileExistsError(f"{target} already exists")
    _retry(lambda: os.replace(path, target), path=path, retries=retries, delay=delay)


def atomic_save(
    workbook: Workbook,
    path: str | Path,
    *,
    retries: int = DEFAULT_RETRIES,
    delay: float = DEFAULT_DELAY,
) -> None:
    """Save *workbook* to *path* without risking the file already there.

    Writes to a temporary file in the same directory, verifies it is a
    readable .xlsx whose every XML part parses, then atomically replaces the
    target. On any failure the temporary file is removed and the original is
    left untouched.

    Raises:
        InvalidFileError: the freshly written file failed validation.
        FileLockedError: the target stayed locked through every retry.
    """
    path = resolve_target(path)
    ensure_writable(path)
    tmp = path.parent / f".{path.stem}.{token_hex(6)}.tmp.xlsx"
    try:
        _retry(lambda: workbook.save(tmp), path=path, retries=retries, delay=delay)
        if not (is_valid_xlsx(tmp) and _is_well_formed(tmp)):
            raise InvalidFileError(
                f"wrote a temporary file for {path} but it failed validation; {path} left unchanged"
            )
        keep_permissions(path, tmp)
        _retry(lambda: os.replace(tmp, path), path=path, retries=retries, delay=delay)
    finally:
        _discard(tmp)


def _discard(tmp: Path) -> None:
    """Remove the temporary file, without letting a failure to do so hide the real error.

    When the save itself fails part-way, Windows can still see the half-written file as open
    until its writer is garbage-collected, and refuses to delete it: a ``PermissionError``
    here would replace whatever actually went wrong. So collect, try once more, and if it
    still won't go, leave it — the error that matters is the one being raised.
    """
    try:
        tmp.unlink(missing_ok=True)
    except OSError:
        gc.collect()
        with suppress(OSError):
            tmp.unlink(missing_ok=True)
