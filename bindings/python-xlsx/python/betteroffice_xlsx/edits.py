"""Typed dictionaries for versioned reads and version-checked edit batches.

Fields keep the camelCase wire names every binding shares, so one request serializes the same
from Python, JavaScript and Rust. Rows and columns are zero-based and ranges inclusive. Sheet ids
come from the current catalog (``sheet:{index}`` standalone, the replica's sheet keys in
collaboration); versions and sheet ids are session-scoped.
"""

from __future__ import annotations

from typing import List, Literal, Optional, TypedDict, Union

EditSource = Literal["host", "agent"]
EditHistory = Literal["separate", "none"]
EditFailureCode = Literal[
    "stale-version",
    "missing-target",
    "content-mismatch",
    "overlapping-steps",
    "locked-target",
    "unsupported",
    "invalid-step",
    "limit-exceeded",
]


class CellPosition(TypedDict):
    row: int
    col: int


class A1Address(TypedDict):
    kind: Literal["a1"]
    a1: str


class RowColAddress(TypedDict):
    kind: Literal["rowCol"]
    start: CellPosition
    end: CellPosition


RangeAddress = Union[A1Address, RowColAddress]


class RangeTarget(TypedDict):
    sheetId: str
    range: RangeAddress


class CellAddress(TypedDict):
    sheetId: str
    row: int
    col: int
    a1: str


class _TaggedValue(TypedDict):
    kind: Literal["empty", "number", "text", "bool", "error"]


class TaggedValue(_TaggedValue, total=False):
    """``{"kind": "number", "value": 1.5}``; errors carry codes such as ``"#DIV/0!"``."""

    value: Union[float, str, bool]


class CellGuard(TypedDict, total=False):
    """Conditions on a cell's pre-batch state; an absent field imposes none."""

    value: TaggedValue
    formula: Optional[str]
    displayText: str


class StepGuard(TypedDict):
    cells: List[List[CellGuard]]


class _Step(TypedDict):
    target: RangeTarget


class _Guarded(_Step, total=False):
    expect: StepGuard


class SetCellInputs(_Guarded):
    """What a user would type, parsed against each cell's current number format."""

    op: Literal["setCellInputs"]
    inputs: List[List[str]]


class SetFormulas(_Guarded):
    """Formula source without the leading ``=``, stored whatever the cell's format."""

    op: Literal["setFormulas"]
    formulas: List[List[str]]


class SetNumberFormat(_Guarded):
    """``format`` is ``"percent"`` and the like or ``{"type": "custom", "pattern": ...}``."""

    op: Literal["setNumberFormat"]
    format: Union[str, dict]


class PatchStyle(_Guarded):
    op: Literal["patchStyle"]
    patch: dict


EditStep = Union[SetCellInputs, SetFormulas, SetNumberFormat, PatchStyle]


class CalculationRequest(TypedDict, total=False):
    """Without ``nowSerial``, volatile functions such as NOW() have no clock."""

    nowSerial: float


class _EditRequest(TypedDict):
    expectVersion: str
    steps: List[EditStep]


class EditRequest(_EditRequest, total=False):
    source: EditSource
    history: EditHistory
    calculation: CalculationRequest


class _Failure(TypedDict):
    code: EditFailureCode
    message: str


class EditFailure(_Failure, total=False):
    stepIndex: int
    conflictingStepIndex: int
    target: RangeTarget


class Refusal(TypedDict):
    ok: Literal[False]
    version: str
    failure: EditFailure


class EditReceipt(TypedDict):
    stepIndex: int
    changed: bool
    target: RangeTarget
    changedCells: List[CellAddress]


class EditPreview(TypedDict):
    stepIndex: int
    target: RangeTarget
    wouldChange: bool
    changedCellCount: int


class EditCalculation(TypedDict):
    changed: List[CellAddress]
    cycleCells: List[CellAddress]
    limitedCells: List[CellAddress]
    truncated: bool


class EditApplication(TypedDict):
    ok: Literal[True]
    baseVersion: str
    version: str
    applied: bool
    source: EditSource
    receipts: List[EditReceipt]
    changedSheets: List[str]
    calculation: EditCalculation


class EditValidation(TypedDict):
    ok: Literal[True]
    baseVersion: str
    wouldApply: bool
    previews: List[EditPreview]


EditResult = Union[EditApplication, Refusal]
ValidationResult = Union[EditValidation, Refusal]


class ReadRequest(TypedDict):
    """An empty ``ranges`` reads only the sheet catalog."""

    ranges: List[RangeTarget]


class SheetEntry(TypedDict):
    sheetId: str
    name: str
    editable: bool


class CellRead(TypedDict):
    a1: str
    value: TaggedValue
    formula: Optional[str]
    displayText: str


class RangeRead(TypedDict):
    target: RangeTarget
    cells: List[List[CellRead]]


class ReadCalculation(TypedDict):
    cycleCells: List[CellAddress]
    limitedCells: List[CellAddress]
    truncated: bool


class CellsRead(TypedDict):
    ok: Literal[True]
    version: str
    sheets: List[SheetEntry]
    ranges: List[RangeRead]
    calculation: ReadCalculation


ReadResult = Union[CellsRead, Refusal]


class _FindRequest(TypedDict):
    text: str


class FindRequest(_FindRequest, total=False):
    """Exact, case-sensitive search over display text; ``limit`` defaults to 100."""

    sheetIds: List[str]
    limit: int


class FindMatch(TypedDict):
    cell: CellAddress
    text: str


class TextFound(TypedDict):
    ok: Literal[True]
    version: str
    matches: List[FindMatch]
    truncated: bool


FindResult = Union[TextFound, Refusal]
