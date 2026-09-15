"""Visual formatting for a Table: fonts, fills, and borders.

A ``TableStyle`` is immutable data — it has no effect on a Table's data,
only its appearance in the .xlsx file. pyhandlexl applies it to the header
row, label column, and data body whenever a table is created or written,
and again whenever it's moved by another table's growth (see "Multiple
named tables on one sheet" in the README) — so a table's look always
extends to cover however much of it currently exists, with no separate
"reformat" step to remember.
"""

from __future__ import annotations

from dataclasses import dataclass


def _is_hex_color(value: str) -> bool:
    return len(value) == 6 and all(c in "0123456789abcdefABCDEF" for c in value)


@dataclass(frozen=True)
class TableStyle:
    """How a Table looks when written: header/label styling, data banding, borders.

    An empty string for ``header_fill``, ``band_fill``, or ``border_color``
    means "none" — no fill, no row banding, or no border at all,
    respectively. Every other field is always meaningful (fonts always have
    a name/size/color, ``header_bold`` is always True or False).

    See ``TableStyle.DEFAULT``, ``.MINIMAL``, and ``.NONE`` for ready-made
    choices — construct your own ``TableStyle(...)`` for anything else.
    """

    header_font_name: str = "Calibri"
    header_font_size: int = 11
    header_font_color: str = "000000"
    header_bold: bool = True
    header_fill: str = "76933C"
    data_font_name: str = "Calibri"
    data_font_size: int = 11
    data_font_color: str = "000000"
    band_fill: str = "F2F2F2"
    border_color: str = "000000"

    def __post_init__(self) -> None:
        for name in ("header_font_name", "data_font_name"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value:
                raise ValueError(f"{name} must be a non-empty str, got {value!r}")
        for name in ("header_font_size", "data_font_size"):
            value = getattr(self, name)
            if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
                raise ValueError(f"{name} must be a positive int, got {value!r}")
        if not isinstance(self.header_bold, bool):
            raise TypeError(f"header_bold must be bool, got {type(self.header_bold).__name__}")
        for name in ("header_font_color", "data_font_color"):
            value = getattr(self, name)
            if not isinstance(value, str) or not _is_hex_color(value):
                raise ValueError(f"{name} must be a 6-digit hex color, got {value!r}")
        for name in ("header_fill", "band_fill", "border_color"):
            value = getattr(self, name)
            if not isinstance(value, str) or (value != "" and not _is_hex_color(value)):
                raise ValueError(f"{name} must be '' or a 6-digit hex color, got {value!r}")


TableStyle.DEFAULT = TableStyle()
TableStyle.MINIMAL = TableStyle(header_fill="", band_fill="")
TableStyle.NONE = TableStyle(header_fill="", header_bold=False, band_fill="", border_color="")
