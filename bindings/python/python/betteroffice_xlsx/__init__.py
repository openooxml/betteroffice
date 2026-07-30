"""Read, recalculate, render, and write XLSX workbooks on the BetterOffice Rust engine."""

from __future__ import annotations

import datetime
import math
import numbers
import os
from typing import Iterable, Iterator, Mapping, Union

from ._betteroffice_xlsx import (
    Calculation,
    CellError,
    CollaborativeStateError,
    History,
    InvalidUpdateError,
    MAX_COLLABORATION_BYTES,
    Mutation,
    NotCollaborativeError,
    Proposal,
    ProposedEdit,
    ParseError,
    Png,
    RangeError,
    RenderError,
    StaleProposalError,
)
from ._betteroffice_xlsx import Workbook as _Workbook
from ._betteroffice_xlsx import XlsxError, __version__

__all__ = [
    "Calculation",
    "CellError",
    "CellValue",
    "CollaborativeStateError",
    "History",
    "InvalidUpdateError",
    "MAX_COLLABORATION_BYTES",
    "Mutation",
    "NotCollaborativeError",
    "Proposal",
    "ProposedEdit",
    "ParseError",
    "Png",
    "RangeError",
    "RenderError",
    "Sheet",
    "SheetKey",
    "StaleProposalError",
    "Workbook",
    "XlsxError",
    "__version__",
]

CellValue = Union[None, bool, float, str, CellError]
SheetKey = Union[int, str]


class Sheet:
    """One sheet, bound to its workbook.

    Assigning writes what a user would type, so a leading ``=`` makes a formula
    and recalculates its dependents. Indexing reads the current value, which for
    an untouched cell is whatever the file cached.
    """

    __slots__ = ("_index", "_workbook")

    def __init__(self, workbook: "Workbook", index: int) -> None:
        self._workbook = workbook
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def name(self) -> str:
        return self._workbook.sheet_names[self._index]

    def __getitem__(self, address: str) -> CellValue:
        return self._workbook.value(self._index, address)

    def __setitem__(self, address: str, value: object) -> None:
        self._workbook.set(self._index, address, value)

    def formula(self, address: str) -> "str | None":
        return self._workbook.formula(self._index, address)

    def render_png(
        self,
        *,
        scale: float = 1.0,
        range: "str | None" = None,
        max_width: "int | None" = None,
        max_height: "int | None" = None,
    ) -> Png:
        return self._workbook.render_png(
            self._index,
            scale=scale,
            range=range,
            max_width=max_width,
            max_height=max_height,
        )

    def __repr__(self) -> str:
        return f"Sheet({self.name!r})"


class Workbook:
    """An XLSX workbook."""

    __slots__ = ("_inner",)

    def __init__(self, inner: _Workbook) -> None:
        self._inner = inner

    @classmethod
    def open(cls, data: "bytes | bytearray | memoryview") -> "Workbook":
        """Open from bytes. Formulas keep the values cached in the file."""
        return cls(_Workbook.open(_as_bytes(data)))

    @classmethod
    def open_path(cls, path: "str | os.PathLike[str]") -> "Workbook":
        """Open from a filesystem path."""
        return cls(_Workbook.open_path(os.fspath(path)))

    @classmethod
    def open_recalculated(
        cls,
        data: "bytes | bytearray | memoryview",
        *,
        now_serial: "float | None" = None,
    ) -> "Workbook":
        """Open from bytes and recalculate every formula up front."""
        return cls(_Workbook.open_recalculated(_as_bytes(data), now_serial=now_serial))

    def recalculate(self, *, now_serial: "float | None" = None) -> Calculation:
        """Recalculate every formula."""
        return self._inner.recalculate(now_serial=now_serial)

    @property
    def sheet_names(self) -> "list[str]":
        return self._inner.sheet_names

    @property
    def sheet_count(self) -> int:
        return self._inner.sheet_count

    def sheet(self, key: SheetKey) -> Sheet:
        """A sheet by name or index. Raises if it does not exist."""
        return Sheet(self, self._inner.sheet_index(key))

    def value(self, sheet: SheetKey, address: str) -> CellValue:
        """The value of a cell, recalculated only if the workbook has been."""
        return self._inner.value(sheet, address)

    def formula(self, sheet: SheetKey, address: str) -> "str | None":
        """The source formula of a cell, without the leading ``=``."""
        return self._inner.formula(sheet, address)

    def set(
        self,
        sheet: SheetKey,
        address: str,
        value: object,
        *,
        now_serial: "float | None" = None,
    ) -> Mutation:
        """Set a cell from what a user would type."""
        return self._inner.set(sheet, address, _as_input(value), now_serial=now_serial)

    def render_png(
        self,
        sheet: SheetKey,
        *,
        scale: float = 1.0,
        range: "str | None" = None,
        max_width: "int | None" = None,
        max_height: "int | None" = None,
    ) -> Png:
        """Render a sheet to PNG."""
        return self._inner.render_png(
            sheet,
            scale=scale,
            range=range,
            max_width=max_width,
            max_height=max_height,
        )

    def save(self) -> bytes:
        """Serialize back to XLSX bytes."""
        return self._inner.save()

    def save_path(self, path: "str | os.PathLike[str]") -> None:
        """Serialize to a filesystem path."""
        self._inner.save_path(os.fspath(path))

    @classmethod
    def open_collaborative(
        cls,
        data: "bytes | bytearray | memoryview",
        *,
        client_id: "int | None" = None,
        recalculate: bool = False,
        now_serial: "float | None" = None,
    ) -> "Workbook":
        """Open a replica that exchanges Yrs updates.

        An omitted ``client_id`` is generated. Explicit IDs must be unique
        among peers because Yrs cannot detect duplicates after authoring.
        """
        return cls(
            _Workbook.open_collaborative(
                _as_bytes(data),
                client_id=client_id,
                recalculate=recalculate,
                now_serial=now_serial,
            )
        )

    @property
    def client_id(self) -> int:
        return self._inner.client_id

    @property
    def is_collaborative(self) -> bool:
        return self._inner.is_collaborative

    def state_vector(self) -> bytes:
        """This replica's state vector, to hand a peer so it can compute a diff."""
        return self._inner.state_vector()

    def state_as_update(self) -> bytes:
        """The whole document as one update, for a peer joining from nothing."""
        return self._inner.state_as_update()

    def diff(self, state_vector: "bytes | bytearray | memoryview") -> bytes:
        """The update carrying everything the peer's state vector is missing."""
        return self._inner.diff(
            _as_bytes(state_vector, max_length=MAX_COLLABORATION_BYTES)
        )

    def apply_update(
        self,
        update: "bytes | bytearray | memoryview",
        *,
        now_serial: "float | None" = None,
    ) -> Mutation:
        return self._inner.apply_update(
            _as_bytes(update, max_length=MAX_COLLABORATION_BYTES),
            now_serial=now_serial,
        )

    @property
    def can_undo(self) -> bool:
        return self._inner.can_undo

    @property
    def can_redo(self) -> bool:
        return self._inner.can_redo

    def history(self) -> History:
        return self._inner.history()

    def undo(self, *, now_serial: "float | None" = None) -> Mutation:
        return self._inner.undo(now_serial=now_serial)

    def redo(self, *, now_serial: "float | None" = None) -> Mutation:
        return self._inner.redo(now_serial=now_serial)

    def set_many(
        self,
        sheet: SheetKey,
        edits: "Mapping[str, object] | Iterable[tuple[str, object]]",
        *,
        now_serial: "float | None" = None,
    ) -> Mutation:
        """Write many cells as one undo step."""
        pairs = edits.items() if hasattr(edits, "items") else edits
        return self._inner.set_many(
            sheet,
            [(address, _as_input(value)) for address, value in pairs],
            now_serial=now_serial,
        )

    def propose(
        self,
        agent_id: str,
        edits: "Iterable[tuple[SheetKey, str, object]]",
        *,
        note: "str | None" = None,
        now_serial: "float | None" = None,
    ) -> Proposal:
        """Stage edits for a human to accept or reject instead of applying them."""
        return self._inner.propose(
            agent_id,
            [(sheet, address, _as_input(value)) for sheet, address, value in edits],
            note=note,
            now_serial=now_serial,
        )

    def proposals(self) -> "list[Proposal]":
        return self._inner.proposals()

    def accept_proposal(
        self,
        proposal_id: str,
        *,
        force: bool = False,
        now_serial: "float | None" = None,
    ) -> Mutation:
        return self._inner.accept_proposal(
            proposal_id, force=force, now_serial=now_serial
        )

    def reject_proposal(self, proposal_id: str) -> bool:
        return self._inner.reject_proposal(proposal_id)

    @property
    def active_sheet(self) -> int:
        return self._inner.active_sheet

    def set_active_sheet(self, sheet: SheetKey) -> None:
        """Select a sheet and persist it as the workbook's active tab."""
        self._inner.set_active_sheet(sheet)

    def merged_ranges(self, sheet: SheetKey, range: str) -> "list[str]":
        return self._inner.merged_ranges(sheet, range)

    def last_calculation(self) -> Calculation:
        return self._inner.last_calculation()

    def set_number_format(
        self,
        sheet: SheetKey,
        range: str,
        format: str,
        *,
        now_serial: "float | None" = None,
    ) -> Mutation:
        """One of automatic, text, number, percent, scientific, currency, date,
        time, or a custom pattern such as ``#,##0.00``."""
        return self._inner.set_number_format(sheet, range, format, now_serial=now_serial)

    def set_style(
        self,
        sheet: SheetKey,
        range: str,
        *,
        bold: "bool | None" = None,
        italic: "bool | None" = None,
        strikethrough: "bool | None" = None,
        font_family: "str | None" = None,
        font_size: "float | None" = None,
        text_color: "str | None" = None,
        fill_color: "str | None" = None,
        horizontal_alignment: "str | None" = None,
        vertical_alignment: "str | None" = None,
        text_wrapping: "str | None" = None,
        now_serial: "float | None" = None,
    ) -> Mutation:
        """Patch styling over a range. Omitted arguments are left as they are."""
        return self._inner.set_style(
            sheet,
            range,
            bold=bold,
            italic=italic,
            strikethrough=strikethrough,
            font_family=font_family,
            font_size=font_size,
            text_color=text_color,
            fill_color=fill_color,
            horizontal_alignment=horizontal_alignment,
            vertical_alignment=vertical_alignment,
            text_wrapping=text_wrapping,
            now_serial=now_serial,
        )

    def __getitem__(self, key: SheetKey) -> Sheet:
        return self.sheet(key)

    def __iter__(self) -> Iterator[Sheet]:
        return (Sheet(self, index) for index in range(self.sheet_count))

    def __len__(self) -> int:
        return self.sheet_count

    def __repr__(self) -> str:
        return f"Workbook(sheets={self.sheet_count})"


def _as_input(value: object) -> str:
    """Coerce a Python value to the string a user would have typed."""
    if value is None:
        return ""
    if value is True:
        return "TRUE"
    if value is False:
        return "FALSE"
    if isinstance(value, str):
        return value
    if isinstance(value, (datetime.date, datetime.time, datetime.timedelta)):
        raise TypeError(
            f"{type(value).__name__} is not supported yet; write the Excel "
            "serial number as a float, or the text you want with str()"
        )
    text = str(value)
    try:
        as_float = float(text)
    except ValueError:
        as_float = None

    if isinstance(value, numbers.Number):
        if as_float is None:
            raise ValueError(
                f"{value!r} serializes to {text!r}, which the engine reads as "
                "text rather than a number; convert it first"
            )
        if not math.isfinite(as_float):
            raise ValueError(
                f"{value!r} is not a finite number and cannot be stored as one; "
                "pass a string if you want the text"
            )
    return text


def _as_bytes(
    data: "bytes | bytearray | memoryview", *, max_length: "int | None" = None
) -> bytes:
    if not isinstance(data, (bytes, bytearray, memoryview)):
        raise TypeError("data must be bytes, bytearray, or memoryview")
    length = memoryview(data).nbytes
    if max_length is not None and length > max_length:
        raise XlsxError(
            f"collaboration payload is {length} bytes, exceeds the "
            f"{max_length}-byte limit"
        )
    return data if isinstance(data, bytes) else bytes(data)
