"""Typed dictionaries for read-only structured exports and their Markdown rendering.

Keys are the engine's camelCase wire fields, shared with the TypeScript contract in
``packages/pptx/src/structuredExport.ts``. Records nested below a slide keep that contract's
shape and are typed loosely here. A ``range`` anchor is an edit-batch text target.
"""

from __future__ import annotations

from typing import Any, Dict, List, Literal, TypedDict, Union

PptxAnchorScope = Literal["session", "snapshot"]
PptxExportFailureCode = Literal["invalid-options", "limit-exceeded", "invalid-content"]
PptxExportSeverity = Literal["info", "warning", "error"]
PptxAnchor = Dict[str, Any]


class PptxExportOptions(TypedDict, total=False):
    includeHiddenSlides: bool
    includeHiddenShapes: bool
    includeNotes: bool
    includeComments: bool
    includeFormatting: bool
    maxBlocks: int
    maxBytes: int


class PptxExportDiagnostic(TypedDict):
    code: str
    severity: PptxExportSeverity
    anchor: Union[PptxAnchor, None]
    message: str


class PptxIncludedContent(TypedDict):
    hiddenSlides: bool
    hiddenShapes: bool
    notes: bool
    comments: bool
    formatting: bool


class PptxStructuredContent(TypedDict):
    schemaVersion: Literal[1]
    anchorScope: PptxAnchorScope
    readingOrder: Literal["shapeTree"]
    included: PptxIncludedContent
    slides: List[Dict[str, Any]]
    diagnostics: List[PptxExportDiagnostic]
    truncated: bool


class PptxMarkdownAnchor(TypedDict):
    marker: str
    anchor: PptxAnchor


class PptxMarkdownContent(TypedDict):
    markdown: str
    anchors: List[PptxMarkdownAnchor]
    diagnostics: List[PptxExportDiagnostic]
    truncated: bool


class PptxExportFailure(TypedDict):
    code: PptxExportFailureCode
    target: Union[PptxAnchor, None]
    message: str


class PptxExportRefusal(TypedDict):
    ok: Literal[False]
    version: str
    failure: PptxExportFailure


class PptxStructuredRead(TypedDict):
    ok: Literal[True]
    version: str
    content: PptxStructuredContent


class PptxMarkdownRead(TypedDict):
    ok: Literal[True]
    version: str
    content: PptxMarkdownContent


PptxStructuredResult = Union[PptxStructuredRead, PptxExportRefusal]
PptxMarkdownResult = Union[PptxMarkdownRead, PptxExportRefusal]
