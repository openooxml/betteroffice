"""Read, edit, and lay out PPTX presentations on the BetterOffice Rust engine."""

from __future__ import annotations

import os
from typing import Iterator, Mapping, Union

from ._betteroffice_pptx import (
    AdjustEdit,
    CollaborativeStateError,
    Deck,
    DisplayList,
    FillEdit,
    InvalidUpdateError,
    MAX_COLLABORATION_BYTES,
    Media,
    NotCollaborativeError,
    Paragraph,
    ParseError,
    RangeError,
    Rect,
    RenderError,
    Shape,
    ShapeEdit,
    Slide,
    SlideEdit,
    Story,
    Stroke,
    StrokeEdit,
    TextEdit,
    TextRun,
    TransformEdit,
)
from ._betteroffice_pptx import Presentation as _Presentation
from ._betteroffice_pptx import PptxError, __version__

__all__ = [
    "AdjustEdit",
    "CollaborativeStateError",
    "Deck",
    "DisplayList",
    "EMU_PER_CENTIMETER",
    "EMU_PER_INCH",
    "EMU_PER_POINT",
    "FillEdit",
    "InvalidUpdateError",
    "MAX_COLLABORATION_BYTES",
    "Media",
    "NotCollaborativeError",
    "Paragraph",
    "ParseError",
    "PptxError",
    "Presentation",
    "RangeError",
    "Rect",
    "RenderError",
    "Shape",
    "ShapeEdit",
    "Slide",
    "SlideEdit",
    "SlideKey",
    "Story",
    "Stroke",
    "StrokeEdit",
    "TextEdit",
    "TextRun",
    "TransformEdit",
    "__version__",
]

EMU_PER_INCH = 914400
EMU_PER_CENTIMETER = 360000
EMU_PER_POINT = 12700

SlideKey = Union[int, str]


class Presentation:
    """A PPTX presentation.

    Geometry is in English Metric Units; text offsets are UTF-16 code units.
    Pinned to the thread that opened it, and must be released there too.
    """

    __slots__ = ("_inner",)

    def __init__(self, inner: _Presentation) -> None:
        self._inner = inner

    @classmethod
    def open(
        cls,
        data: "bytes | bytearray | memoryview",
        *,
        limits: "Mapping[str, int] | None" = None,
    ) -> "Presentation":
        """Open from bytes. `limits` overrides individual parser bounds."""
        return cls(_Presentation.open(_as_bytes(data), limits=_as_limits(limits)))

    @classmethod
    def open_path(
        cls,
        path: "str | os.PathLike[str]",
        *,
        limits: "Mapping[str, int] | None" = None,
    ) -> "Presentation":
        """Open from a filesystem path."""
        return cls(_Presentation.open_path(os.fspath(path), limits=_as_limits(limits)))

    @classmethod
    def open_collaborative(
        cls,
        data: "bytes | bytearray | memoryview",
        *,
        client_id: "int | None" = None,
        limits: "Mapping[str, int] | None" = None,
    ) -> "Presentation":
        """Open a replica that exchanges Yrs updates.

        An omitted ``client_id`` is generated. Explicit IDs must be unique
        among peers because Yrs cannot detect duplicates after authoring.
        """
        return cls(
            _Presentation.open_collaborative(
                _as_bytes(data), client_id=client_id, limits=_as_limits(limits)
            )
        )

    @property
    def client_id(self) -> int:
        return self._inner.client_id

    @property
    def is_collaborative(self) -> bool:
        """True only for ``open_collaborative`` decks, which alone may sync."""
        return self._inner.is_collaborative

    @property
    def is_edited(self) -> bool:
        """True once the engine has accepted an edit since the deck opened."""
        return self._inner.is_edited

    @property
    def author(self) -> str:
        """Recorded on every edit this replica makes."""
        return self._inner.author

    @author.setter
    def author(self, author: str) -> None:
        self._inner.author = author

    @property
    def origin(self) -> str:
        """One of local, agent, remote, or system.

        Only ``local`` edits enter this replica's undo stack.
        """
        return self._inner.origin

    @origin.setter
    def origin(self, origin: str) -> None:
        self._inner.origin = origin

    @property
    def slide_count(self) -> int:
        return self._inner.slide_count

    @property
    def slide_ids(self) -> "list[str]":
        return self._inner.slide_ids

    @property
    def width_emu(self) -> int:
        return self._inner.width_emu

    @property
    def height_emu(self) -> int:
        return self._inner.height_emu

    @property
    def layouts(self) -> "list[str]":
        """Layout part paths, as ``insert_slide`` expects them."""
        return self._inner.layouts

    def snapshot(self) -> Deck:
        """The whole deck as plain data, as of now."""
        return self._inner.snapshot()

    def slide(self, slide: SlideKey) -> Slide:
        """One slide by ID or index."""
        return self._inner.slide(slide)

    def story(self, story_id: str) -> Story:
        """One text flow by ID. Raises ``KeyError`` if it does not exist."""
        return self._inner.story(story_id)

    def media(self) -> "list[Media]":
        """Every embedded binary part, such as slide images."""
        return self._inner.media()

    def insert_slide(self, index: int, *, layout: "str | None" = None) -> SlideEdit:
        """Insert a blank slide at ``index``."""
        return self._inner.insert_slide(index, layout=layout)

    def delete_slide(self, slide: SlideKey) -> SlideEdit:
        return self._inner.delete_slide(slide)

    def move_slide(self, slide: SlideKey, index: int) -> SlideEdit:
        return self._inner.move_slide(slide, index)

    def add_text_box(
        self,
        slide: SlideKey,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        text: str = "",
        name: "str | None" = None,
        bold: "bool | None" = None,
        italic: "bool | None" = None,
        font_size: "float | None" = None,
        color: "str | None" = None,
        font_family: "str | None" = None,
        underline: "str | None" = None,
    ) -> ShapeEdit:
        """Add a text box. Geometry is in EMU."""
        return self._inner.add_text_box(
            slide,
            x=x,
            y=y,
            width=width,
            height=height,
            text=text,
            name=name,
            bold=bold,
            italic=italic,
            font_size=font_size,
            color=color,
            font_family=font_family,
            underline=underline,
        )

    def add_shape(
        self,
        slide: SlideKey,
        geometry: str,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        fill: "str | None" = None,
        name: "str | None" = None,
    ) -> ShapeEdit:
        """Add a preset shape such as ``rect``, ``ellipse``, or ``roundRect``."""
        return self._inner.add_shape(
            slide, geometry, x=x, y=y, width=width, height=height, fill=fill, name=name
        )

    def set_shape_fill(
        self, slide: SlideKey, shape_id: str, color: "str | None"
    ) -> FillEdit:
        """Set a solid fill. ``None`` clears it."""
        return self._inner.set_shape_fill(slide, shape_id, color)

    def set_shape_stroke(
        self,
        slide: SlideKey,
        shape_id: str,
        *,
        color: "str | None" = None,
        width_pt: "float | None" = None,
    ) -> StrokeEdit:
        """Set the outline. Omitting both arguments clears it."""
        return self._inner.set_shape_stroke(
            slide, shape_id, color=color, width_pt=width_pt
        )

    def set_shape_adjust(
        self, slide: SlideKey, shape_id: str, adjustments: "Mapping[str, float]"
    ) -> AdjustEdit:
        """Set preset geometry guides. Values are clamped into legal range."""
        return self._inner.set_shape_adjust(slide, shape_id, dict(adjustments))

    def remove_shape(self, slide: SlideKey, shape_id: str) -> ShapeEdit:
        return self._inner.remove_shape(slide, shape_id)

    def move_shape(
        self, slide: SlideKey, shape_id: str, x: int, y: int
    ) -> TransformEdit:
        return self._inner.move_shape(slide, shape_id, x, y)

    def resize_shape(
        self, slide: SlideKey, shape_id: str, width: int, height: int
    ) -> TransformEdit:
        return self._inner.resize_shape(slide, shape_id, width, height)

    def insert_text(
        self,
        story_id: str,
        index: int,
        text: str,
        *,
        bold: "bool | None" = None,
        italic: "bool | None" = None,
        font_size: "float | None" = None,
        color: "str | None" = None,
        font_family: "str | None" = None,
        underline: "str | None" = None,
    ) -> TextEdit:
        """Insert text at a UTF-16 offset in a story."""
        return self._inner.insert_text(
            story_id,
            index,
            text,
            bold=bold,
            italic=italic,
            font_size=font_size,
            color=color,
            font_family=font_family,
            underline=underline,
        )

    def delete_text(self, story_id: str, start: int, end: int) -> TextEdit:
        """Delete a UTF-16 range; the receipt carries the removed text."""
        return self._inner.delete_text(story_id, start, end)

    def format_text(
        self,
        story_id: str,
        start: int,
        end: int,
        *,
        bold: "bool | None" = None,
        italic: "bool | None" = None,
        font_size: "float | None" = None,
        color: "str | None" = None,
        font_family: "str | None" = None,
        underline: "str | None" = None,
    ) -> TextEdit:
        """Patch run styling over a UTF-16 range. Omitted arguments are left as they are."""
        return self._inner.format_text(
            story_id,
            start,
            end,
            bold=bold,
            italic=italic,
            font_size=font_size,
            color=color,
            font_family=font_family,
            underline=underline,
        )

    def insert_paragraph_break(self, story_id: str, index: int) -> TextEdit:
        return self._inner.insert_paragraph_break(story_id, index)

    def register_font(
        self,
        family: str,
        data: "bytes | bytearray | memoryview",
        *,
        bold: bool = False,
        italic: bool = False,
    ) -> int:
        """Register a face for layout. No face is embedded in the wheel."""
        return self._inner.register_font(
            family, _as_bytes(data), bold=bold, italic=italic
        )

    def render_slide(self, slide: SlideKey) -> DisplayList:
        """Lay a slide out into the renderer's display list."""
        return self._inner.render_slide(slide)

    def save(self) -> bytes:
        """Serialize back to PPTX bytes, edits included."""
        return self._inner.save()

    def save_path(self, path: "str | os.PathLike[str]") -> None:
        """Serialize to a filesystem path, edits included."""
        self._inner.save_path(os.fspath(path))

    def state_vector(self) -> bytes:
        """This replica's state vector, to hand a peer so it can compute a diff.

        Raises ``NotCollaborativeError`` unless the deck is collaborative.
        """
        return self._inner.state_vector()

    def state_as_update(self) -> bytes:
        """The whole document as one update, for a peer joining from nothing."""
        return self._inner.state_as_update()

    def diff(self, state_vector: "bytes | bytearray | memoryview") -> bytes:
        """The update carrying everything the peer's state vector is missing."""
        return self._inner.diff(
            _as_bytes(state_vector, max_length=MAX_COLLABORATION_BYTES)
        )

    def apply_update(self, update: "bytes | bytearray | memoryview") -> Deck:
        """Merge a peer's update and return the deck it produced."""
        return self._inner.apply_update(
            _as_bytes(update, max_length=MAX_COLLABORATION_BYTES)
        )

    @property
    def can_undo(self) -> bool:
        return self._inner.can_undo

    @property
    def can_redo(self) -> bool:
        return self._inner.can_redo

    def undo(self) -> bool:
        return self._inner.undo()

    def redo(self) -> bool:
        return self._inner.redo()

    def add_undo_barrier(self) -> None:
        """Close the current undo step so the next edit starts a new one."""
        self._inner.add_undo_barrier()

    def __getitem__(self, slide: SlideKey) -> Slide:
        return self._inner.slide(slide)

    def __iter__(self) -> Iterator[Slide]:
        return (self._inner.slide(index) for index in range(self.slide_count))

    def __len__(self) -> int:
        return self.slide_count

    def __repr__(self) -> str:
        return f"Presentation(slides={self.slide_count})"


def _as_limits(limits: "Mapping[str, int] | None") -> "dict[str, int] | None":
    return None if limits is None else dict(limits)


def _as_bytes(
    data: "bytes | bytearray | memoryview", *, max_length: "int | None" = None
) -> bytes:
    if not isinstance(data, (bytes, bytearray, memoryview)):
        raise TypeError("data must be bytes, bytearray, or memoryview")
    length = memoryview(data).nbytes
    if max_length is not None and length > max_length:
        raise PptxError(
            f"collaboration payload is {length} bytes, exceeds the "
            f"{max_length}-byte limit"
        )
    return data if isinstance(data, bytes) else bytes(data)
