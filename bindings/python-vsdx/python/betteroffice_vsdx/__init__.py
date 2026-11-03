"""Read Visio VSDX diagrams on the BetterOffice Rust engine."""

from typing import TypedDict

from ._betteroffice_vsdx import Cell, Connect, Png, Diagram, Page, ParseError, RangeError, RenderError, Shape, VsdxError, __version__
from ._betteroffice_vsdx import decode_font_face

class PngTarget(TypedDict):
    page: int


__all__ = ["Png", "PngTarget", "Cell", "Connect", "Diagram", "Page", "ParseError", "RangeError", "RenderError", "Shape", "VsdxError", "__version__", "decode_font_face"]
