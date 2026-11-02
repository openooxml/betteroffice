from __future__ import annotations
from os import PathLike
from typing import final

def decode_font_face(
    data: bytes, codec: str, *, dict: bytes | None = None, max_output: int | None = None
) -> tuple[bytes, bytes | None]:
    """Decode a font pack face and return font bytes and an optional transformed payload."""
    ...

class VsdxError(Exception): ...
class ParseError(VsdxError): ...
class RangeError(VsdxError): ...
class RenderError(VsdxError): ...

@final
class Cell:
    name: str
    formula: str | None
    value: str | None
    unit: str | None

@final
class Connect:
    from_sheet: int
    from_cell: str | None
    from_part: int | None
    to_sheet: int
    to_cell: str | None
    to_part: int | None

@final
class Shape:
    id: int
    name: str | None
    text: str | None
    pin_x: float | None
    pin_y: float | None
    width: float | None
    height: float | None
    cells: list[Cell]
    children: list[Shape]

@final
class Page:
    id: int
    name: str | None
    source_part_path: str
    shapes: list[Shape]
    connects: list[Connect]

@final
class Diagram:
    @staticmethod
    def open(data: bytes | bytearray | memoryview) -> Diagram: ...
    @staticmethod
    def open_path(path: str | PathLike[str]) -> Diagram: ...
    @property
    def pages(self) -> list[Page]: ...
    def __len__(self) -> int: ...
