"""Shared console rendering for ``Table.show`` and ``grid.show``.

Internal module — not part of the public API.
"""

from __future__ import annotations


def _text(value: object) -> str:
    """*value* as one line of display text.

    ``None`` is blank, everything else is ``str()``. A newline or carriage return in the
    value — a data cell can genuinely hold one — is made visible (``\\n``) rather than left
    as a real line break, which would otherwise split a bordered row across lines and wreck
    the layout.
    """
    if value is None:
        return ""
    return str(value).replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")


def render_box(rows: list[list[object]], *, header_rows: int = 0) -> str:
    """*rows* as a bordered, column-aligned block: a rule above the first row and below the
    last, and (with *header_rows*) another rule after that many rows — unless they're all
    there is, since the closing rule already covers it. Every row must be the same length.

    Returns the finished text, with no trailing newline; ``""`` for no rows at all.
    """
    if not rows:
        return ""
    width = len(rows[0])
    text_rows = [[_text(value) for value in row] for row in rows]
    col_widths = [max(len(row[c]) for row in text_rows) for c in range(width)]
    rule = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"

    last = len(text_rows) - 1
    lines = [rule]
    for i, row in enumerate(text_rows):
        cells = (cell.ljust(w) for cell, w in zip(row, col_widths, strict=True))
        lines.append("| " + " | ".join(cells) + " |")
        if header_rows and i == header_rows - 1 and i != last:
            lines.append(rule)
    lines.append(rule)
    return "\n".join(lines)
