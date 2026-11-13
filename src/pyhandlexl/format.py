"""Cell-level formatting: ``CellFormat`` for one cell's look, ``Styled`` to carry it in a grid.

A ``CellFormat`` changes how a single cell looks in the .xlsx file and nothing else — never
its value. On a :class:`~pyhandlexl.Table` it is attached to a (row label, column header)
pair with ``set_format``, so it stays on the same cell however the table is edited or
moved. In a grid (the list of rows that ``read_sheet`` returns) it travels with the value,
wrapped in ``Styled``, so every ``grid`` helper moves the format along with the value.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

Align = Literal["left", "center", "right"]
_ALIGNS = ("left", "center", "right")


def _is_hex_color(value: str) -> bool:
    return len(value) == 6 and all(c in "0123456789abcdefABCDEF" for c in value)


@dataclass(frozen=True)
class CellFormat:
    """How one cell looks. Every field is optional: ``None`` means "leave that part alone".

    Only the parts you give are applied, on top of the cell's usual look (a table's
    :class:`~pyhandlexl.TableStyle` banding, say), so ``CellFormat(bold=True)`` makes a cell
    bold without touching its colours.

    Args:
        bold: bold text on or off.
        italic: italic text on or off.
        font_color: text colour as a 6-digit hex string (``"C00000"``).
        fill: background colour as a 6-digit hex string, or ``""`` for no fill at all.
        number_format: an Excel number format code (``"0.00"``, ``"#,##0"``, ``"0%"``,
            ``"yyyy-mm-dd"``). It changes how a number or date is *shown*, not its value.
        align: horizontal alignment — ``"left"``, ``"center"`` or ``"right"``.
        wrap: wrap long text onto several lines inside the cell, on or off. Excel grows
            the row to fit the wrapped text unless the row has a fixed height.
    """

    bold: bool | None = None
    italic: bool | None = None
    font_color: str | None = None
    fill: str | None = None
    number_format: str | None = None
    align: Align | None = None
    wrap: bool | None = None

    def __post_init__(self) -> None:
        for name in ("bold", "italic", "wrap"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{name} must be bool or None, got {type(value).__name__}")
        if self.font_color is not None and (
            not isinstance(self.font_color, str) or not _is_hex_color(self.font_color)
        ):
            raise ValueError(f"font_color must be a 6-digit hex color, got {self.font_color!r}")
        if self.fill is not None and (
            not isinstance(self.fill, str) or (self.fill != "" and not _is_hex_color(self.fill))
        ):
            raise ValueError(f"fill must be '' or a 6-digit hex color, got {self.fill!r}")
        if self.number_format is not None and (
            not isinstance(self.number_format, str) or not self.number_format
        ):
            raise ValueError(f"number_format must be a non-empty str, got {self.number_format!r}")
        if self.align is not None and self.align not in _ALIGNS:
            raise ValueError(f"align must be one of {_ALIGNS}, got {self.align!r}")

    def is_empty(self) -> bool:
        """Whether this format changes nothing (every field is ``None``)."""
        return self == _EMPTY


_EMPTY = CellFormat()


@dataclass(frozen=True)
class Styled:
    """A grid value together with its :class:`CellFormat`.

    ``read_sheet(path, formats=True)`` returns formatted cells this way, and
    ``write_sheet`` / ``append_rows`` accept them. The ``grid`` helpers treat a ``Styled``
    like any other value, so inserting, deleting or transposing moves the format with the
    value it belongs to.
    """

    value: object
    format: CellFormat

    def __post_init__(self) -> None:
        if not isinstance(self.format, CellFormat):
            raise TypeError(f"format must be a CellFormat, got {type(self.format).__name__}")
        if isinstance(self.value, Styled):
            raise TypeError("value must not itself be Styled")


def unwrap(value: object) -> object:
    """The plain value of *value*: ``Styled(v, f)`` gives ``v``, anything else is itself."""
    return value.value if isinstance(value, Styled) else value
