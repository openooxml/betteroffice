"""Read, recalculate, render, and write XLSX workbooks on the BetterOffice Rust engine."""

from __future__ import annotations

import datetime
import math
import numbers
import os
from typing import Iterator, Union

from ._betteroffice_xlsx import Calculation, CellError, ParseError, Png, RangeError, RenderError
from ._betteroffice_xlsx import Workbook as _Workbook
from ._betteroffice_xlsx import XlsxError, __version__

__all__ = [
    "Calculation",
    "CellError",
    "CellValue",
    "ParseError",
    "Png",
    "RangeError",
    "RenderError",
    "Sheet",
    "SheetKey",
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
        return cls(_Workbook.open(bytes(data)))

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
        return cls(_Workbook.open_recalculated(bytes(data), now_serial=now_serial))

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
    ) -> bool:
        """Set a cell from what a user would type. Returns whether it changed."""
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
