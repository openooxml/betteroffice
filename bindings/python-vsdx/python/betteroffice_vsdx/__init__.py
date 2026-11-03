"""Read Visio VSDX diagrams on the BetterOffice Rust engine."""

from typing import TypedDict

from ._betteroffice_vsdx import Cell, Connect, Png, Diagram, Page, ParseError, RangeError, RenderError, Shape, VsdxError, __version__
from ._betteroffice_vsdx import decode_font_face

class PageTarget(TypedDict):
    page: int


__all__ = ["Png", "PageTarget", "Cell", "Connect", "Diagram", "Page", "ParseError", "RangeError", "RenderError", "Shape", "VsdxError", "__version__", "decode_font_face"]
