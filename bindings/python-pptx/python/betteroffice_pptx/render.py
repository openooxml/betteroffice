"""Renderer and package inspection wire types."""

from __future__ import annotations

from typing import Any, Dict, List, Literal, TypedDict, Union


class ShapeHit(TypedDict):
    kind: Literal["shape"]
    shapeId: str


class TextHit(TypedDict):
    kind: Literal["text"]
    shapeId: str
    storyId: str
    position: int


HitTestResult = Union[ShapeHit, TextHit]


class ProposalTextChange(TypedDict):
    storyId: str
    start: int
    end: int
    kind: Literal["insertion", "deletion"]


class _SlideDisplayListOptional(TypedDict, total=False):
    background: Dict[str, Any]


class SlideDisplayList(_SlideDisplayListOptional):
    contractVersion: int
    width: float
    height: float
    primitives: List[Dict[str, Any]]


class ProposalDiffSlide(TypedDict):
    proposal: Dict[str, Any]
    snapshot: Dict[str, Any]
    textChanges: List[ProposalTextChange]
    frame: SlideDisplayList


class MediaSummary(TypedDict):
    partPath: str
    contentType: str
    byteLength: int


class MediaContent(TypedDict):
    partPath: str
    contentType: str
    bytes: str


class _PresentationInspectionOptional(TypedDict, total=False):
    diagramDrawings: List[Dict[str, Any]]
    tableStyles: Dict[str, Any]
    commentAuthors: List[Dict[str, Any]]
    comments: List[Dict[str, Any]]
    commentFlavor: Literal["legacy", "modern"]
    shapeElements: Literal["withConnectors"]


class PresentationInspection(_PresentationInspectionOptional):
    presentation: Dict[str, Any]
    slides: List[Dict[str, Any]]
    layouts: List[Dict[str, Any]]
    masters: List[Dict[str, Any]]
    themes: List[Dict[str, Any]]
    charts: List[Dict[str, Any]]
    media: List[Union[MediaSummary, MediaContent]]
    relationships: Dict[str, List[Dict[str, Any]]]
