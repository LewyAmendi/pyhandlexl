"""Tests for the internal _safety module."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import pyhandlexl._safety as safety
from pyhandlexl._safety import atomic_save, safe_delete, safe_load
from pyhandlexl.errors import FileLockedError, InvalidFileError


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    """Make retries instant so the suite stays fast."""
    monkeypatch.setattr(safety, "sleep", lambda _seconds: None)


def _make_workbook(path: Path, a1: str = "hello") -> None:
    wb = Workbook()
    wb.active["A1"] = a1
    wb.save(path)


class TestRetry:
    def test_succeeds_after_transient_failures(self):
        attempts = []

        def flaky():
            attempts.append(1)
            if len(attempts) < 3:
                raise PermissionError("locked")
            return "done"

        assert safety._retry(flaky, path=Path("x.xlsx"), retries=5) == "done"
        assert len(attempts) == 3

    def test_raises_file_locked_when_exhausted(self):
        def always_locked():
            raise PermissionError("still locked")

        with pytest.raises(FileLockedError):
            safety._retry(always_locked, path=Path("x.xlsx"), retries=3)

    def test_non_transient_error_is_not_retried(self):
        attempts = []

        def boom():
            attempts.append(1)
            raise ValueError("not a lock")

        with pytest.raises(ValueError):
            safety._retry(boom, path=Path("x.xlsx"), retries=5)
        assert len(attempts) == 1


class TestSafeDelete:
    def test_deletes_the_file(self, tmp_path):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p)
        safe_delete(p)
        assert not p.exists()

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            safe_delete(tmp_path / "nope.xlsx")

    def test_locked_file_retries_then_raises(self, tmp_path, monkeypatch):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p)

        def locked_unlink(self):
            raise PermissionError("open in Excel")

        monkeypatch.setattr(Path, "unlink", locked_unlink)
        with pytest.raises(FileLockedError):
            safe_delete(p)


class TestSafeLoad:
    def test_loads_existing_workbook(self, tmp_path):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p)
        assert safe_load(p).active["A1"].value == "hello"

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            safe_load(tmp_path / "nope.xlsx")

    def test_invalid_file_raises(self, tmp_path):
        p = tmp_path / "bad.xlsx"
        p.write_text("not a workbook")
        with pytest.raises(InvalidFileError):
            safe_load(p)


class TestAtomicSave:
    def test_creates_new_file(self, tmp_path):
        p = tmp_path / "new.xlsx"
        atomic_save(Workbook(), p)
        assert p.is_file()

    def test_overwrites_existing_content(self, tmp_path):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p, a1="old")
        wb = Workbook()
        wb.active["A1"] = "new"
        atomic_save(wb, p)
        assert load_workbook(p).active["A1"].value == "new"

    def test_no_temp_file_left_on_success(self, tmp_path):
        p = tmp_path / "wb.xlsx"
        atomic_save(Workbook(), p)
        assert {f.name for f in tmp_path.iterdir()} == {"wb.xlsx"}

    def test_original_preserved_when_validation_fails(self, tmp_path, monkeypatch):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p, a1="keep me")

        monkeypatch.setattr(safety, "is_valid_xlsx", lambda _path: False)
        doomed = Workbook()
        doomed.active["A1"] = "should not land"
        with pytest.raises(InvalidFileError):
            atomic_save(doomed, p)

        assert load_workbook(p).active["A1"].value == "keep me"
        assert {f.name for f in tmp_path.iterdir()} == {"wb.xlsx"}

    def test_locked_target_raises_file_locked_and_cleans_up(self, tmp_path, monkeypatch):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p)

        def locked_replace(src, dst):
            raise PermissionError("target open in Excel")

        monkeypatch.setattr(safety.os, "replace", locked_replace)
        with pytest.raises(FileLockedError):
            atomic_save(Workbook(), p)

        assert {f.name for f in tmp_path.iterdir()} == {"wb.xlsx"}


class TestWellFormedCheck:
    @staticmethod
    def _archive(tmp_path: Path, part: bytes) -> Path:
        from zipfile import ZipFile

        path = tmp_path / "x.xlsx"
        with ZipFile(path, "w") as z:
            z.writestr("xl/worksheets/sheet1.xml", part)
        return path

    def test_a_real_workbook_is_well_formed(self, tmp_path):
        p = tmp_path / "wb.xlsx"
        _make_workbook(p)
        assert safety._is_well_formed(p)

    @pytest.mark.parametrize(
        "part",
        [
            b'<?xml version="1.0"?><a><b></a>',  # mismatched tag
            b'<?xml version="1.0"?><a><b>',  # truncated
            b'<?xml version="1.0"?><x:a/>',  # a namespace prefix nobody declared
            b'<?xml version="1.0"?><a>\x00</a>',  # a character XML can't hold
            b'<?xml version="1.0"?><a>&undefined;</a>',
            b"",
        ],
        ids=["mismatched", "truncated", "unbound-prefix", "illegal-char", "entity", "empty"],
    )
    def test_a_broken_part_is_refused(self, tmp_path, part):
        assert not safety._is_well_formed(self._archive(tmp_path, part))

    def test_a_well_formed_part_with_namespaces_is_accepted(self, tmp_path):
        part = b'<?xml version="1.0"?><x:a xmlns:x="urn:x"><x:b y="1"/></x:a>'
        assert safety._is_well_formed(self._archive(tmp_path, part))

    def test_a_file_that_is_not_a_zip_is_refused(self, tmp_path):
        path = tmp_path / "x.xlsx"
        path.write_bytes(b"not a zip")
        assert not safety._is_well_formed(path)


class TestCleanupNeverHidesTheRealError:
    def test_a_failed_save_raises_its_own_error_not_a_cleanup_error(self, tmp_path):
        target = tmp_path / "wb.xlsx"
        _make_workbook(target)

        class Failing(Workbook):
            handle = None

            def save(self, filename):  # type: ignore[override]
                # like a writer that dies part-way: the half-written file is still open
                Failing.handle = open(filename, "wb")  # noqa: SIM115
                raise RuntimeError("the writer failed")

        try:
            with pytest.raises(RuntimeError, match="the writer failed"):
                atomic_save(Failing(), target)
        finally:
            if Failing.handle is not None:
                Failing.handle.close()
        load_workbook(target).close()  # the original is untouched and still opens

    def test_discard_ignores_a_file_it_cannot_delete(self, tmp_path, monkeypatch):
        stuck = tmp_path / "stuck.tmp.xlsx"
        stuck.write_bytes(b"x")

        def refuse(self, missing_ok=False):
            raise PermissionError("still open")

        monkeypatch.setattr(Path, "unlink", refuse)
        safety._discard(stuck)  # must not raise
