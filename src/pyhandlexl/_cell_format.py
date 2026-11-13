"""Internal: applying a ``CellFormat`` to an openpyxl cell, reading one back, and storing them.

Not part of the public API.
"""

from __future__ import annotations

import json
from typing import Any

from openpyxl.styles import Alignment, Font, PatternFill

from pyhandlexl.format import CellFormat
from pyhandlexl.validate import MAX_CELL_CHARS


def apply_format(cell: Any, fmt: CellFormat) -> None:
    """Apply the parts of *fmt* that are set to *cell*, leaving the rest as they are."""
    if fmt.bold is not None or fmt.italic is not None or fmt.font_color is not None:
        font = cell.font
        cell.font = Font(
            name=font.name,
            size=font.size,
            bold=font.bold if fmt.bold is None else fmt.bold,
            italic=font.italic if fmt.italic is None else fmt.italic,
            underline=font.underline,
            color=font.color if fmt.font_color is None else fmt.font_color,
        )
    if fmt.fill is not None:
        cell.fill = (
            PatternFill(start_color=fmt.fill, end_color=fmt.fill, fill_type="solid")
            if fmt.fill
            else PatternFill(fill_type=None)
        )
    if fmt.number_format is not None:
        cell.number_format = fmt.number_format
    if fmt.align is not None or fmt.wrap is not None:
        current = cell.alignment  # keep whatever else the cell's alignment already says
        cell.alignment = Alignment(
            horizontal=current.horizontal if fmt.align is None else fmt.align,
            vertical=current.vertical,
            wrap_text=current.wrap_text if fmt.wrap is None else fmt.wrap,
            indent=current.indent,
            text_rotation=current.textRotation,
            shrink_to_fit=current.shrinkToFit,
        )


def _rgb(color: Any) -> str | None:
    rgb = getattr(color, "rgb", None) if color is not None else None
    if isinstance(rgb, str) and len(rgb) == 8 and all(c in "0123456789abcdefABCDEF" for c in rgb):
        return rgb[2:].upper()
    return None


def read_format(cell: Any) -> CellFormat:
    """The ``CellFormat`` a cell visibly carries (empty if it looks like a default cell)."""
    font = cell.font
    color = _rgb(font.color)
    fill = cell.fill
    fill_color = _rgb(fill.fgColor) if fill is not None and fill.fill_type == "solid" else None
    number_format = cell.number_format
    align = cell.alignment.horizontal
    return CellFormat(
        bold=True if font.bold else None,
        italic=True if font.italic else None,
        font_color=color if color not in (None, "000000") else None,
        fill=fill_color,
        number_format=number_format if number_format != "General" else None,
        align=align if align in ("left", "center", "right") else None,
        wrap=True if cell.alignment.wrap_text else None,
    )


# ------------------------------------------------------------ storage in the schema


def _fmt_to_list(fmt: CellFormat) -> list[Any]:
    return [fmt.bold, fmt.italic, fmt.font_color, fmt.fill, fmt.number_format, fmt.align, fmt.wrap]


def _fmt_from_list(items: list[Any]) -> CellFormat:
    return CellFormat(*items)


def formats_to_json(formats: dict[tuple[int, int], CellFormat]) -> list[str]:
    """*formats* (keyed by (data row index, data column index)) as one or more JSON strings,
    each short enough to fit a single cell. Empty if there are no formats."""
    chunks: list[str] = []
    table: list[list[Any]] = []
    index: dict[CellFormat, int] = {}
    cells: list[list[int]] = []
    size = 0

    def flush() -> None:
        nonlocal table, index, cells, size
        if cells:
            chunks.append(json.dumps({"f": table, "c": cells}, separators=(",", ":")))
        table, index, cells, size = [], {}, [], 0

    for (i, j), fmt in sorted(formats.items()):
        extra = 0
        if fmt not in index:
            extra = len(json.dumps(_fmt_to_list(fmt), separators=(",", ":"))) + 1
        cell_size = len(f"[{i},{j},{len(table)}],") + extra
        if cells and size + cell_size > MAX_CELL_CHARS - 64:
            flush()
            extra = len(json.dumps(_fmt_to_list(fmt), separators=(",", ":"))) + 1
            cell_size = len(f"[{i},{j},0],") + extra
        if fmt not in index:
            index[fmt] = len(table)
            table.append(_fmt_to_list(fmt))
        cells.append([i, j, index[fmt]])
        size += cell_size
    flush()
    return chunks


def dims_to_json(
    widths: dict[int, float],
    heights: dict[int, float],
    autos: dict[int, float],
    fit: tuple[bool, int] | None,
) -> list[str]:
    """Column widths (keyed by local column), row heights (by data row), the widths the library
    last fitted itself, and the fit settings, as JSON strings that each fit one cell."""
    chunks: list[str] = []
    if fit is not None:
        chunks.append(json.dumps({"s": [int(fit[0]), fit[1]]}, separators=(",", ":")))
    cur: dict[str, list[list[float]]] = {}
    size = 0
    for tag, source in (("w", widths), ("h", heights), ("a", autos)):
        for key, value in sorted(source.items()):
            piece = len(f"[{key},{value}],")
            if cur and size + piece > MAX_CELL_CHARS - 64:
                chunks.append(json.dumps(cur, separators=(",", ":")))
                cur, size = {}, 0
            cur.setdefault(tag, []).append([key, value])
            size += piece
    if cur:
        chunks.append(json.dumps(cur, separators=(",", ":")))
    return chunks


def dims_from_json(
    chunks: list[str],
) -> tuple[dict[int, float], dict[int, float], dict[int, float], tuple[bool, int] | None]:
    """The inverse of :func:`dims_to_json`."""
    widths: dict[int, float] = {}
    heights: dict[int, float] = {}
    autos: dict[int, float] = {}
    fit: tuple[bool, int] | None = None
    for chunk in chunks:
        data = json.loads(chunk)
        if "s" in data:
            fit = (bool(data["s"][0]), int(data["s"][1]))
        for tag, target in (("w", widths), ("h", heights), ("a", autos)):
            for key, value in data.get(tag, ()):
                target[int(key)] = float(value)
    return widths, heights, autos, fit


def formats_from_json(chunks: list[str]) -> dict[tuple[int, int], CellFormat]:
    """The inverse of :func:`formats_to_json`."""
    out: dict[tuple[int, int], CellFormat] = {}
    for chunk in chunks:
        data = json.loads(chunk)
        if "f" not in data:  # a chunk of another kind (see dims_from_json)
            continue
        table = [_fmt_from_list(items) for items in data["f"]]
        for i, j, k in data["c"]:
            out[(i, j)] = table[k]
    return out
