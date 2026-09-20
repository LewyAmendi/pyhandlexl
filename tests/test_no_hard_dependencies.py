"""numpy and pandas are optional: the library must work, and never import them, without them."""

from __future__ import annotations

import subprocess
import sys
import textwrap


def run(code: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-c", textwrap.dedent(code)], capture_output=True, text=True
    )


def test_everything_works_with_numpy_and_pandas_missing(tmp_path):
    path = tmp_path / "wb.xlsx"
    result = run(
        f"""
        import sys
        sys.modules["numpy"] = None   # an import of either now fails, as if not installed
        sys.modules["pandas"] = None
        import pyhandlexl as p

        p.create_workbook({str(path)!r})
        p.write_sheet({str(path)!r}, [[1, float("nan"), "x"]])
        p.create_sheet({str(path)!r}, "D")
        p.Table(data=[[1, None]], row_labels=["a"], column_headers=["x", "y"], name="T").create(
            {str(path)!r}, sheet="D"
        )
        assert p.read_sheet({str(path)!r}, "Sheet") == [[1, None, "x"]]
        assert p.Table.read({str(path)!r}, "T").data.rows == [[1, None]]
        p.check_cell_value(float("nan"))
        """
    )
    assert result.returncode == 0, result.stderr


def test_the_library_imports_neither_at_module_level():
    # (openpyxl itself imports numpy when it is installed, so sys.modules can't tell whose
    # import it was: read the source instead.)
    import ast
    from pathlib import Path

    import pyhandlexl

    offenders = []
    for path in Path(pyhandlexl.__file__).parent.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names = [node.module]
            if any(name.split(".")[0] in ("numpy", "pandas") for name in names):
                offenders.append(f"{path.name}:{node.lineno}")
    assert not offenders, offenders
