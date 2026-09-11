"""Read, edit, lay out, and render DOCX documents on the BetterOffice Rust engine."""

from __future__ import annotations

import os
from typing import Any, Iterator, Mapping, Union

from ._betteroffice_docx import (
    DisplayList,
    Edit,
    EditError,
    HeaderFooter,
    Layout,
    LayoutError,
    MAX_PIXMAP_DIM,
    MAX_PIXMAP_PIXELS,
    Paragraph,
    ParseError,
    Png,
    RenderError,
    Section,
    Structure,
    Table,
    TableCell,
    TableRow,
    TextRun,
    UnsupportedEditError,
)
from ._betteroffice_docx import Document as _Document
from ._betteroffice_docx import DocxError, __version__

__all__ = [
    "DisplayList",
    "Document",
    "DocxError",
    "Edit",
    "EditError",
    "HeaderFooter",
    "Layout",
    "LayoutError",
    "MAX_PIXMAP_DIM",
    "MAX_PIXMAP_PIXELS",
    "Paragraph",
    "ParagraphKey",
    "ParseError",
    "Png",
    "RenderError",
    "Section",
    "Structure",
    "TWIPS_PER_INCH",
    "TWIPS_PER_POINT",
    "Table",
    "TableCell",
    "TableRow",
    "TextRun",
    "UnsupportedEditError",
    "__version__",
]

TWIPS_PER_INCH = 1440
TWIPS_PER_POINT = 20

ParagraphKey = Union[int, str]


class Document:
    """A DOCX document.

    Page geometry is in twips; text offsets are UTF-16 code units. Reads return
    values, not live views, so read again after an edit.
    """

    __slots__ = ("_inner",)

    def __init__(self, inner: _Document) -> None:
        self._inner = inner

    @classmethod
    def open(
        cls,
        data: "bytes | bytearray | memoryview",
        *,
        limits: "Mapping[str, int] | None" = None,
    ) -> "Document":
        """Open from bytes. `limits` overrides individual parser bounds."""
        return cls(_Document.open(_as_bytes(data), limits=_as_limits(limits)))

    @classmethod
    def open_path(
        cls,
        path: "str | os.PathLike[str]",
        *,
        limits: "Mapping[str, int] | None" = None,
    ) -> "Document":
        """Open from a filesystem path."""
        return cls(_Document.open_path(os.fspath(path), limits=_as_limits(limits)))

    @property
    def author(self) -> str:
        """Recorded on every edit this document makes."""
        return self._inner.author

    @author.setter
    def author(self, author: str) -> None:
        self._inner.author = author

    @property
    def origin(self) -> str:
        """One of local, agent, remote, or system."""
        return self._inner.origin

    @origin.setter
    def origin(self, origin: str) -> None:
        self._inner.origin = origin

    @property
    def timestamp(self) -> str:
        """The ISO-8601 clock reading edits and saves record.

        The engine has no clock, so this defaults to the epoch and output stays
        byte-reproducible until you set one.
        """
        return self._inner.timestamp

    @timestamp.setter
    def timestamp(self, timestamp: str) -> None:
        self._inner.timestamp = timestamp

    @property
    def paragraph_ids(self) -> "list[str | None]":
        """Body paragraph IDs in document order.

        A paragraph Word never stamped with a ``w14:paraId`` reads as ``None``
        and cannot be addressed by ID.
        """
        return self._inner.paragraph_ids

    @property
    def warnings(self) -> "list[str]":
        """What the parser could not represent. Parsing still succeeded."""
        return self._inner.warnings

    @property
    def template_variables(self) -> "list[str]":
        """``{{name}}`` placeholders the parser found in body text."""
        return self._inner.template_variables

    @property
    def text(self) -> str:
        """Every body paragraph's text, one per line."""
        return self._inner.text

    def structure(self) -> Structure:
        """How much of each kind the document holds."""
        return self._inner.structure()

    def paragraphs(self) -> "list[Paragraph]":
        """Every body paragraph, table cells and content controls included."""
        return self._inner.paragraphs()

    def tables(self) -> "list[Table]":
        return self._inner.tables()

    def sections(self) -> "list[Section]":
        return self._inner.sections()

    def headers(self) -> "list[HeaderFooter]":
        return self._inner.headers()

    def footers(self) -> "list[HeaderFooter]":
        return self._inner.footers()

    def paragraph(self, paragraph: ParagraphKey) -> Paragraph:
        """One paragraph by ID or by index into the body."""
        return self._inner.paragraph(paragraph)

    def replace_text(self, para_id: str, text: str) -> Edit:
        """Rewrite one paragraph's text, keeping its style and run formatting.

        Raises ``KeyError`` for an unknown ID and ``UnsupportedEditError`` for a
        paragraph the engine cannot rebuild from a single run.
        """
        return self._inner.replace_text(para_id, text)

    def layout(self, input: "str | Mapping[str, Any]") -> Layout:
        """Paginate a ``{"measured": [...], "options": {...}}`` envelope.

        The engine paginates blocks that were measured already; it does not
        measure them, so the envelope carries the metrics.
        """
        return self._inner.layout(input if isinstance(input, str) else dict(input))

    def register_font(
        self,
        family: str,
        data: "bytes | bytearray | memoryview",
        *,
        bold: bool = False,
        italic: bool = False,
    ) -> int:
        """Register a face for rasterization. No face is embedded in the wheel."""
        return self._inner.register_font(family, _as_bytes(data), bold=bold, italic=italic)

    def register_image(
        self,
        rel_id: str,
        data: "bytes | bytearray | memoryview",
        *,
        scope: str = "body",
        part: "str | None" = None,
    ) -> None:
        """Supply the bytes behind one unresolved relationship id.

        ``scope`` is body, header_footer, footnotes, or endnotes; the
        header_footer scope needs the owning part's relationship id in ``part``.
        """
        self._inner.register_image(rel_id, _as_bytes(data), scope=scope, part=part)

    def render_png(self, display_list: DisplayList, page: int = 0) -> Png:
        """Rasterize one display-list page to deterministic PNG bytes."""
        return self._inner.render_png(display_list, page)

    def save(
        self,
        *,
        now: "str | None" = None,
        update_modified_date: bool = False,
        modified_by: "str | None" = None,
    ) -> bytes:
        """Serialize back to DOCX bytes, edits included."""
        return self._inner.save(
            now=now, update_modified_date=update_modified_date, modified_by=modified_by
        )

    def save_path(
        self,
        path: "str | os.PathLike[str]",
        *,
        now: "str | None" = None,
        update_modified_date: bool = False,
        modified_by: "str | None" = None,
    ) -> None:
        """Serialize to a filesystem path."""
        self._inner.save_path(
            os.fspath(path),
            now=now,
            update_modified_date=update_modified_date,
            modified_by=modified_by,
        )

    def __getitem__(self, paragraph: ParagraphKey) -> Paragraph:
        return self._inner.paragraph(paragraph)

    def __iter__(self) -> Iterator[Paragraph]:
        return iter(self._inner.paragraphs())

    def __len__(self) -> int:
        return len(self._inner)

    def __repr__(self) -> str:
        return repr(self._inner)


def _as_limits(limits: "Mapping[str, int] | None") -> "dict[str, int] | None":
    return None if limits is None else dict(limits)


def _as_bytes(data: "bytes | bytearray | memoryview") -> bytes:
    if not isinstance(data, (bytes, bytearray, memoryview)):
        raise TypeError("data must be bytes, bytearray, or memoryview")
    return data if isinstance(data, bytes) else bytes(data)
