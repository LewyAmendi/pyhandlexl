"""Three-way merge of a table's state, so concurrent writers don't lose each other's work.

Internal module — the public pieces (``MergeReport``, ``MergeConflictWarning``)
are re-exported from the package root.

Three states are in play:

* **base**   — the table as this ``Table`` object last saw it on disk.
* **ours**   — what the object holds now, after the caller's edits.
* **theirs** — what is on disk right now (someone else wrote since *base*).

Rows and columns are identified by label / header, which a table keeps unique.
The one operation a label can't express is a rename, so the ``Table`` tracks
those (``row_origin`` / ``col_origin``: our label -> the base label it started
as, or ``None`` for something we added). Everything else is read off a diff.

Policy: changes that don't overlap merge silently; where both sides changed the
same thing differently, **ours wins** and the overwrite is reported as a
conflict.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta

from pyhandlexl.column_type import ColumnType
from pyhandlexl.style import TableStyle
from pyhandlexl.validate import normalize_newlines

_MISSING = object()
_TOLERANCE = timedelta(milliseconds=1)  # Excel's date/time storage rounds sub-millisecond digits


@dataclass
class Snapshot:
    """One table's full state, independent of any file."""

    row_labels: list[str]
    column_headers: list[str]
    data: list[list[object]]  # data[i][j]: row i, column j
    corner: str
    style: TableStyle
    column_types: list[ColumnType]  # aligned with column_headers

    def copy(self) -> Snapshot:
        return Snapshot(
            list(self.row_labels),
            list(self.column_headers),
            [list(row) for row in self.data],
            self.corner,
            self.style,
            list(self.column_types),
        )


@dataclass(frozen=True)
class MergeReport:
    """What ``Table.write`` reconciled with changes another writer made meanwhile.

    Available afterwards as ``t.last_merge`` (``None`` if nothing had changed on
    disk). Everything here is already reflected in the table's data; the
    ``conflicts`` are the places where yours overwrote theirs, which is also
    reported as a ``MergeConflictWarning``.
    """

    rows_added: tuple[str, ...] = ()  # rows another writer added, now in your table
    rows_removed: tuple[str, ...] = ()  # rows another writer deleted, gone from yours
    columns_added: tuple[str, ...] = ()
    columns_removed: tuple[str, ...] = ()
    cells_updated: int = 0  # cells you already held whose value came from the other writer
    conflicts: tuple[str, ...] = ()  # where yours overwrote theirs


# --------------------------------------------------------------- comparison


def _kind(value: object) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, (datetime, date)):
        return "datetime"
    if isinstance(value, time):
        return "time"
    if isinstance(value, timedelta):
        return "timedelta"
    return type(value).__name__


def _as_datetime(value: date) -> datetime:
    return value if isinstance(value, datetime) else datetime(value.year, value.month, value.day)


def _seconds(value: time) -> float:
    return value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6


def same_value(a: object, b: object) -> bool:
    """Strict cell equality that survives what Excel does to values on a round trip.

    ``True`` is not ``1``; ``5`` is ``5.0``; a ``date`` is the ``datetime`` at
    its midnight; date/time values compare to the millisecond.
    """
    if a is b:
        return True
    if a is None or b is None:
        return False
    kind = _kind(a)
    if kind != _kind(b):
        return False
    if kind == "number":
        return float(a) == float(b)
    if kind == "datetime":
        return abs(_as_datetime(a) - _as_datetime(b)) < _TOLERANCE
    if kind == "time":
        return abs(_seconds(a) - _seconds(b)) < _TOLERANCE.total_seconds()
    if kind == "timedelta":
        return abs(a - b) < _TOLERANCE
    if kind == "str":  # a carriage return is stored as a newline
        return a == b or normalize_newlines(a) == normalize_newlines(b)
    return a == b


def same_state(a: Snapshot, b: Snapshot) -> bool:
    """Whether two snapshots hold the same table (labels, data, corner, style, types)."""
    return (
        a.row_labels == b.row_labels
        and a.column_headers == b.column_headers
        and a.corner == b.corner
        and a.style == b.style
        and a.column_types == b.column_types
        and all(
            same_value(x, y)
            for row_a, row_b in zip(a.data, b.data, strict=True)
            for x, y in zip(row_a, row_b, strict=True)
        )
    )


def _unique(snapshot: Snapshot) -> bool:
    return len(set(snapshot.row_labels)) == len(snapshot.row_labels) and len(
        set(snapshot.column_headers)
    ) == len(snapshot.column_headers)


def _show(value: object) -> str:
    text = repr(value)
    return text if len(text) <= 30 else text[:27] + "..."


# ------------------------------------------------------------------- an axis


@dataclass
class _Item:
    """One row (or column) of the merged table and where each side has it."""

    label: str  # its label in the merged table
    ours: str | None  # its label on our side, if we have it
    theirs: str | None
    base: str | None  # the base label it descends from, if any


def _anchor(
    items: list[_Item],
    ours_labels: list[str],
    k: int,
    origin: dict[str, str | None],
    dropped: set[str],
) -> int:
    """Where to insert our k-th item.

    An append (nothing that came from the base follows it on our side) goes to the
    end, after whatever the other writer appended, so sequential appenders stay in
    order. Otherwise it goes right after the nearest preceding one of ours that
    survived, or at the very start if it was first, or at the end if everything
    before it is gone.
    """
    if not any(
        origin[later] is not None and later not in dropped for later in ours_labels[k + 1 :]
    ):
        return len(items)
    if k == 0:
        return 0
    for j in range(k - 1, -1, -1):
        for index, item in enumerate(items):
            if item.ours == ours_labels[j]:
                return index + 1
    return len(items)


def _merge_axis(
    kind: str,
    base_labels: list[str],
    ours_labels: list[str],
    origin: dict[str, str | None],
    theirs_labels: list[str],
    ours_edited: Callable[[str], bool],
    theirs_edited: Callable[[str], bool],
    conflicts: list[str],
) -> tuple[list[_Item], list[str]]:
    """Merge the row (or column) labels. Returns the merged items and the labels
    (on our side) that the other writer deleted and we didn't touch."""
    base_set, theirs_set = set(base_labels), set(theirs_labels)
    ours_of_base = {b: o for o, b in origin.items() if b is not None}

    items: list[_Item] = []
    for t in theirs_labels:
        if t in base_set:
            if t in ours_of_base:
                items.append(_Item(ours_of_base[t], ours_of_base[t], t, t))
            elif theirs_edited(t):
                conflicts.append(
                    f"{kind} {t!r}: deleted by you but edited by another writer "
                    "- their edits were discarded"
                )
        else:
            items.append(_Item(t, None, t, None))

    # We renamed a base row onto a label another writer just added: ours wins.
    by_label: dict[str, _Item] = {}
    for item in list(items):
        other = by_label.get(item.label)
        if other is None:
            by_label[item.label] = item
            continue
        keep, lose = (item, other) if other.ours is None else (other, item)
        items.remove(lose)
        by_label[item.label] = keep
        conflicts.append(
            f"{kind} {item.label!r}: another writer added a different {kind} with a name "
            "you renamed one to - yours won"
        )

    dropped: list[str] = []
    resurrect: set[str] = set()
    for b in base_labels:
        if b in ours_of_base and b not in theirs_set:  # the other writer deleted it
            o = ours_of_base[b]
            if ours_edited(b):
                resurrect.add(o)
                conflicts.append(
                    f"{kind} {o!r}: deleted by another writer but edited by you "
                    "- yours was restored"
                )
            else:
                dropped.append(o)

    for k, o in enumerate(ours_labels):
        if origin[o] is not None and o not in resurrect:
            continue
        existing = next((item for item in items if item.label == o), None)
        if existing is not None:
            if existing.ours is None and existing.base is None:  # both sides added it
                existing.ours = o
                continue
            items.remove(existing)
            conflicts.append(
                f"{kind} {o!r}: another writer has a different {kind} with this name "
                "- yours replaced it"
            )
        items.insert(
            _anchor(items, ours_labels, k, origin, set(dropped)), _Item(o, o, None, origin[o])
        )

    return items, dropped


# --------------------------------------------------------------------- merge

NOT_UNIQUE = (
    "row labels or column headers are not unique, so the tables can't be told apart "
    "row by row - yours replaced what was on disk"
)


def merge(
    base: Snapshot,
    ours: Snapshot,
    row_origin: dict[str, str | None],
    col_origin: dict[str, str | None],
    theirs: Snapshot,
) -> tuple[Snapshot, MergeReport]:
    """Merge *theirs* into *ours* relative to *base*. Ours wins where they collide."""
    if not (_unique(base) and _unique(ours) and _unique(theirs)):
        return ours.copy(), MergeReport(conflicts=(NOT_UNIQUE,))

    b_row = {label: i for i, label in enumerate(base.row_labels)}
    b_col = {label: j for j, label in enumerate(base.column_headers)}
    o_row = {label: i for i, label in enumerate(ours.row_labels)}
    o_col = {label: j for j, label in enumerate(ours.column_headers)}
    t_row = {label: i for i, label in enumerate(theirs.row_labels)}
    t_col = {label: j for j, label in enumerate(theirs.column_headers)}
    ours_of_base_row = {b: o for o, b in row_origin.items() if b is not None}
    ours_of_base_col = {b: o for o, b in col_origin.items() if b is not None}

    # ---- "did this side change that row/column?" -- only base cells count, plus a
    # rename (ours) or a filled-in cell in a column/row the other side added.
    def row_ours_edited(b: str) -> bool:
        o = ours_of_base_row[b]
        if o != b:
            return True
        return any(
            not same_value(ours.data[o_row[o]][o_col[oc]], base.data[b_row[b]][b_col[c]])
            for c, oc in ours_of_base_col.items()
        )

    def row_theirs_edited(b: str) -> bool:
        row = theirs.data[t_row[b]]
        if any(
            not same_value(row[t_col[c]], base.data[b_row[b]][b_col[c]])
            for c in base.column_headers
            if c in t_col
        ):
            return True
        return any(row[j] is not None for c, j in t_col.items() if c not in b_col)

    def col_ours_edited(c: str) -> bool:
        oc = ours_of_base_col[c]
        if oc != c:
            return True
        return any(
            not same_value(ours.data[o_row[orow]][o_col[oc]], base.data[b_row[r]][b_col[c]])
            for r, orow in ours_of_base_row.items()
        )

    def col_theirs_edited(c: str) -> bool:
        j = t_col[c]
        if any(
            not same_value(theirs.data[t_row[r]][j], base.data[b_row[r]][b_col[c]])
            for r in base.row_labels
            if r in t_row
        ):
            return True
        return any(theirs.data[i][j] is not None for r, i in t_row.items() if r not in b_row)

    conflicts: list[str] = []
    rows, rows_dropped = _merge_axis(
        "row",
        base.row_labels,
        ours.row_labels,
        row_origin,
        theirs.row_labels,
        row_ours_edited,
        row_theirs_edited,
        conflicts,
    )
    cols, cols_dropped = _merge_axis(
        "column",
        base.column_headers,
        ours.column_headers,
        col_origin,
        theirs.column_headers,
        col_ours_edited,
        col_theirs_edited,
        conflicts,
    )

    if not cols:  # both sides dropped different columns, leaving none: a table needs one
        return ours.copy(), MergeReport(
            conflicts=(
                "another writer's changes would have left the table with no columns "
                "- yours replaced them",
            )
        )

    # ---- cells
    cells_updated = 0
    data: list[list[object]] = []
    for r in rows:
        line: list[object] = []
        for c in cols:
            b = _MISSING
            if r.base is not None and c.base is not None:
                b = base.data[b_row[r.base]][b_col[c.base]]
            o = _MISSING
            if r.ours is not None and c.ours is not None:
                o = ours.data[o_row[r.ours]][o_col[c.ours]]
            t = _MISSING
            if r.theirs is not None and c.theirs is not None:
                t = theirs.data[t_row[r.theirs]][t_col[c.theirs]]

            if o is _MISSING:
                value = None if t is _MISSING else t
            elif t is _MISSING:
                value = o
            elif b is not _MISSING and same_value(o, b):  # we left it alone
                value = t
            else:
                value = o
                if (b is _MISSING or not same_value(t, b)) and not same_value(t, o):
                    conflicts.append(
                        f"cell ({r.label!r}, {c.label!r}): yours {_show(o)} overwrote another "
                        f"writer's {_show(t)}"
                    )
            if o is not _MISSING and not same_value(value, o):
                cells_updated += 1
            line.append(value)
        data.append(line)

    # ---- corner, style, column types: ours only if we changed it since base
    corner = theirs.corner
    if ours.corner != base.corner:
        corner = ours.corner
        if theirs.corner not in (base.corner, ours.corner):
            conflicts.append(f"corner: yours {ours.corner!r} overwrote {theirs.corner!r}")

    style = theirs.style
    if ours.style != base.style:
        style = ours.style
        if theirs.style not in (base.style, ours.style):
            conflicts.append("style: yours overwrote another writer's style change")

    types: list[ColumnType] = []
    for c in cols:
        ours_type = ours.column_types[o_col[c.ours]] if c.ours is not None else None
        theirs_type = theirs.column_types[t_col[c.theirs]] if c.theirs is not None else None
        base_type = base.column_types[b_col[c.base]] if c.base is not None else None
        if ours_type is None:
            chosen = theirs_type
        elif theirs_type is None:
            chosen = ours_type
        elif base_type is None:  # both added the column
            chosen = ours_type
            if theirs_type != ours_type:
                conflicts.append(f"column {c.label!r}: yours type {ours_type.value} won")
        elif ours_type != base_type:
            chosen = ours_type
            if theirs_type not in (base_type, ours_type):
                conflicts.append(
                    f"column {c.label!r}: yours type {ours_type.value} overwrote "
                    f"{theirs_type.value}"
                )
        else:
            chosen = theirs_type
        types.append(chosen if chosen is not None else ColumnType.ANY)

    merged = Snapshot(
        [r.label for r in rows],
        [c.label for c in cols],
        data,
        corner,
        style,
        types,
    )
    report = MergeReport(
        rows_added=tuple(r.label for r in rows if r.ours is None),
        rows_removed=tuple(rows_dropped),
        columns_added=tuple(c.label for c in cols if c.ours is None),
        columns_removed=tuple(cols_dropped),
        cells_updated=cells_updated,
        conflicts=tuple(conflicts),
    )
    return merged, report
