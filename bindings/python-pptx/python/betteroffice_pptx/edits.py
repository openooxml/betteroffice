"""Typed dictionaries for versioned reads and version-checked edit batches.

Keys are the engine's camelCase wire fields, shared with the TypeScript contract in
``packages/pptx/src/edits.ts``. Text offsets are story-local UTF-16 code units; a story
reads as its paragraphs joined by ``\\n``.
"""

from __future__ import annotations

from typing import Any, Dict, List, Literal, TypedDict, Union

PptxEditSource = Literal["host", "agent"]
PptxEditHistory = Literal["separate", "none"]
PptxEditFailureCode = Literal[
    "stale-version",
    "missing-target",
    "ambiguous-target",
    "content-mismatch",
    "overlapping-steps",
    "unsupported",
    "invalid-step",
    "limit-exceeded",
]


class PptxSlideTarget(TypedDict):
    slideId: str


class PptxShapeTarget(PptxSlideTarget):
    shapeId: str


class PptxStoryTarget(PptxShapeTarget):
    storyId: str


class PptxTextRange(PptxStoryTarget):
    start: int
    end: int


class PptxRangeTarget(PptxTextRange):
    kind: Literal["range"]


class PptxSearchTarget(TypedDict):
    """The one exact, case-sensitive, paragraph-local match of ``text``."""

    kind: Literal["search"]
    within: PptxStoryTarget
    text: str


PptxTextTarget = Union[PptxRangeTarget, PptxSearchTarget]


class PptxTextGuard(TypedDict):
    text: str


class _Step(TypedDict, total=False):
    expect: Dict[str, Any]


class PptxInsertTextStep(_Step):
    op: Literal["insertText"]
    target: PptxTextTarget
    at: Literal["start", "end"]
    text: str


class PptxReplaceTextStep(_Step):
    op: Literal["replaceText"]
    target: PptxTextTarget
    text: str


class PptxDeleteTextStep(_Step):
    op: Literal["deleteText"]
    target: PptxTextTarget


class PptxFormatTextStep(_Step):
    op: Literal["formatText"]
    target: PptxTextTarget
    patch: Dict[str, Any]


class PptxParagraphAlignmentStep(_Step):
    op: Literal["setParagraphAlignment"]
    target: PptxTextTarget
    alignment: Union[str, None]


class PptxSlideNotesStep(_Step):
    op: Literal["setSlideNotes"]
    target: PptxSlideTarget
    text: str


class PptxShapeRectStep(_Step):
    op: Literal["setShapeRect"]
    target: PptxShapeTarget
    rect: Dict[str, int]


class PptxShapeFillStep(_Step):
    op: Literal["setShapeFill"]
    target: PptxShapeTarget
    color: Union[str, None]


class PptxShapeStrokeStep(_Step):
    op: Literal["setShapeStroke"]
    target: PptxShapeTarget
    stroke: Dict[str, Any]


PptxEditStep = Union[
    PptxInsertTextStep,
    PptxReplaceTextStep,
    PptxDeleteTextStep,
    PptxFormatTextStep,
    PptxParagraphAlignmentStep,
    PptxSlideNotesStep,
    PptxShapeRectStep,
    PptxShapeFillStep,
    PptxShapeStrokeStep,
]


class _EditRequest(TypedDict):
    expectVersion: str
    steps: List[PptxEditStep]


class PptxEditRequest(_EditRequest, total=False):
    source: PptxEditSource
    history: PptxEditHistory


class _EditFailure(TypedDict):
    code: PptxEditFailureCode
    message: str


class PptxEditFailure(_EditFailure, total=False):
    stepIndex: int
    conflictingStepIndex: int
    target: Dict[str, Any]


class PptxEditRefusal(TypedDict):
    ok: Literal[False]
    version: str
    failure: PptxEditFailure


class PptxEditReceipt(TypedDict):
    stepIndex: int
    changed: bool
    target: Dict[str, Any]


class PptxEditPreview(TypedDict):
    stepIndex: int
    target: Dict[str, Any]
    wouldChange: bool


class PptxEditApplication(TypedDict):
    ok: Literal[True]
    baseVersion: str
    version: str
    applied: bool
    source: PptxEditSource
    changedSlides: List[str]
    changedStories: List[str]
    receipts: List[PptxEditReceipt]


class PptxEditValidation(TypedDict):
    ok: Literal[True]
    baseVersion: str
    wouldApply: bool
    previews: List[PptxEditPreview]


PptxEditResult = Union[PptxEditApplication, PptxEditRefusal]
PptxValidationResult = Union[PptxEditValidation, PptxEditRefusal]


class PptxReadRequest(TypedDict, total=False):
    slideIds: List[str]


class _TextField(TypedDict):
    start: int
    end: int


class PptxTextField(_TextField, total=False):
    fieldType: str


class PptxParagraphText(TypedDict):
    paragraphId: str
    start: int
    end: int
    lineBreaks: List[int]
    fields: List[PptxTextField]
    editable: bool


class PptxStoryText(PptxStoryTarget):
    text: str
    paragraphs: List[PptxParagraphText]


class PptxReadContent(TypedDict):
    ok: Literal[True]
    version: str
    slides: List[Dict[str, Any]]
    stories: List[PptxStoryText]


PptxReadResult = Union[PptxReadContent, PptxEditRefusal]


class _FindScope(TypedDict):
    slideId: str


class PptxFindScope(_FindScope, total=False):
    shapeId: str
    storyId: str


class _FindRequest(TypedDict):
    text: str


class PptxFindRequest(_FindRequest, total=False):
    within: PptxFindScope
    limit: int


class PptxFindMatch(TypedDict):
    text: str
    range: PptxTextRange


class PptxFindContent(TypedDict):
    ok: Literal[True]
    version: str
    matches: List[PptxFindMatch]
    truncated: bool


PptxFindResult = Union[PptxFindContent, PptxEditRefusal]
