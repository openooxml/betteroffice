"""Typed dictionaries for the read-only structured export and its Markdown projection.

Fields keep the camelCase wire names every binding shares. Export never recalculates: formula
results are the stored values (``calculation.policy == "asStored"``, freshness
``"unverified"``). Anchors name positions by sheet id, index and name and A1 in the exported
version or snapshot; none follows later row, column or sheet edits. A cell or range anchor's
``{"sheetId": anchor["sheet"]["sheetId"], "range": {"kind": "a1", "a1": anchor["a1"]}}`` is its
edit-batch target at the exported version.
"""

from __future__ import annotations

from typing import List, Literal, Optional, TypedDict, Union


class _ExportScope(TypedDict):
    sheet: int


class ExportScope(_ExportScope, total=False):
    """A zero-based sheet index and at most one rectangular A1 range (default: used range)."""

    range: str


class ExportOptions(TypedDict, total=False):
    """Hidden sheets, rows, columns and names are excluded unless asked for."""

    scope: List[ExportScope]
    includeHiddenSheets: bool
    includeHiddenRows: bool
    includeHiddenColumns: bool
    includeDefinedNames: bool
    includeHiddenNames: bool
    maxCells: int
    maxBytes: int


class MarkdownOptions(TypedDict, total=False):
    """Grid bounds per sheet and in total; empty grid positions count toward ``maxCells``."""

    maxRows: int
    maxColumns: int
    maxCells: int
    maxBytes: int


class SheetIdentity(TypedDict):
    """``sheetId`` is the exporting session's catalog id, ``sheet:{index}`` for bytes."""

    sheetId: str
    index: int
    name: str


class Anchor(TypedDict, total=False):
    """``kind`` is ``sheet``, ``cell``, ``range``, ``definedName`` or ``sourcePart``."""

    kind: Literal["sheet", "cell", "range", "definedName", "sourcePart"]
    sheet: SheetIdentity
    a1: str
    name: str
    localSheet: Optional[SheetIdentity]
    ordinal: int
    part: str
    partSha256: str
    path: List[int]


class Diagnostic(TypedDict):
    code: str
    severity: Literal["info", "warning", "error"]
    anchor: Optional[Anchor]
    message: str


class ExportCell(TypedDict):
    id: str
    anchor: Anchor
    value: dict
    formula: Optional[str]
    displayText: str
    numberFormat: str
    formulaResult: Optional[
        Literal["unverified", "missing", "uncertain", "cycle", "limited"]
    ]
    merge: Optional[dict]


class ExportSheet(TypedDict):
    id: str
    anchor: Anchor
    kind: Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet", "other"]
    visibility: Literal["visible", "hidden", "veryHidden", "unknown"]
    usedRange: Optional[str]
    selectedRange: Optional[str]
    hiddenRows: List[str]
    hiddenColumns: List[str]
    merges: List[dict]
    tables: List[dict]
    hyperlinks: List[dict]
    objects: List[dict]
    cells: List[ExportCell]
    source: Optional[dict]
    truncated: bool


class StructuredContent(TypedDict):
    schemaVersion: Literal[1]
    anchorScope: Literal["session", "snapshot"]
    dateSystem: Literal["1900", "1904"]
    calculation: dict
    included: dict
    sheets: List[ExportSheet]
    definedNames: List[dict]
    diagnostics: List[Diagnostic]
    truncated: bool


class MarkdownAnchor(TypedDict):
    marker: str
    anchor: Anchor


class MarkdownContent(TypedDict):
    markdown: str
    anchors: List[MarkdownAnchor]
    diagnostics: List[Diagnostic]
    truncated: bool


class ExportFailure(TypedDict):
    code: Literal["invalid-options", "invalid-scope", "limit-exceeded"]
    target: Optional[Anchor]
    message: str


class _Exported(TypedDict):
    ok: Literal[True]
    version: str


class StructuredExport(_Exported):
    content: StructuredContent


class MarkdownExport(_Exported):
    content: MarkdownContent


class ExportRefusal(TypedDict):
    ok: Literal[False]
    version: str
    failure: ExportFailure


StructuredResult = Union[StructuredExport, ExportRefusal]
MarkdownResult = Union[MarkdownExport, ExportRefusal]
