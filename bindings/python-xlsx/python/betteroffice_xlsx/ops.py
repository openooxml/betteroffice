"""Typed dictionaries for the raw edit ops :meth:`Workbook.apply_ops` takes.

Each op is the JSON object the TypeScript ``applyOps`` and the Rust ``Op`` share, tagged by
``type``. Field names are exactly the wire names, which mix camelCase and snake_case
(``refers_to``, ``external_target``, ``edit_as``). Sheets are zero-based indices, rows and
columns zero-based, and ranges inclusive. A field typed ``Optional`` may be left out.
"""

from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional, TypedDict, Union

from .edits import TaggedValue

HorizontalAlignment = Literal["left", "center", "right"]
VerticalAlignment = Literal["top", "middle", "bottom"]
TextWrapping = Literal["overflow", "wrap", "clip"]
BorderPreset = Literal[
    "all", "inner", "horizontal", "vertical", "outer", "left", "top", "right", "bottom", "none"
]
BorderStyle = Literal["solid", "dashed", "dotted", "double"]
StyleProperty = Literal[
    "bold",
    "italic",
    "strikethrough",
    "fontFamily",
    "fontSize",
    "textColor",
    "fillColor",
    "borders",
    "horizontalAlignment",
    "verticalAlignment",
    "textWrapping",
]


class _CellPoint(TypedDict):
    row: int
    col: int


class CellPoint(_CellPoint, total=False):
    """A cell; ``abs_row`` and ``abs_col`` carry ``$`` anchoring and default to false."""

    abs_row: bool
    abs_col: bool


class CellRange(TypedDict):
    start: CellPoint
    end: CellPoint


class CellState(TypedDict, total=False):
    """What ``setCell`` writes; an absent ``value`` is empty."""

    value: TaggedValue
    formula: Optional[str]
    style: Optional[int]


class FreezePane(TypedDict):
    rows: int
    cols: int
    top_left: CellPoint


class _Hyperlink(TypedDict):
    range: CellRange


class Hyperlink(_Hyperlink, total=False):
    external_target: Optional[str]
    location: Optional[str]
    tooltip: Optional[str]
    display: Optional[str]


class DifferentialFormat(TypedDict, total=False):
    fontColor: str
    bold: bool
    italic: bool
    fill: str


class AnchorCell(TypedDict):
    col: int
    colOff: int
    row: int
    rowOff: int


class AnchorExtent(TypedDict):
    cx: int
    cy: int


class AnchorPos(TypedDict):
    x: int
    y: int


TwoCellAnchor = TypedDict(
    "TwoCellAnchor",
    {
        "kind": Literal["twoCell"],
        "from": AnchorCell,
        "to": AnchorCell,
        "edit_as": Literal["twoCell", "oneCell", "absolute"],
    },
)
OneCellAnchor = TypedDict(
    "OneCellAnchor",
    {"kind": Literal["oneCell"], "from": AnchorCell, "extent": AnchorExtent},
)


class AbsoluteAnchor(TypedDict):
    kind: Literal["absolute"]
    pos: AnchorPos
    extent: AnchorExtent


ChartAnchor = Union[TwoCellAnchor, OneCellAnchor, AbsoluteAnchor]


class BorderPatch(TypedDict, total=False):
    preset: BorderPreset
    style: BorderStyle
    color: str


class StylePatch(TypedDict, total=False):
    """Styling to patch; omitted properties are left as they are, ``clear`` resets them."""

    bold: bool
    italic: bool
    strikethrough: bool
    fontFamily: str
    fontSize: float
    textColor: str
    fillColor: str
    border: BorderPatch
    horizontalAlignment: HorizontalAlignment
    verticalAlignment: VerticalAlignment
    textWrapping: TextWrapping
    clear: List[StyleProperty]


class PresetNumberFormat(TypedDict):
    type: Literal[
        "automatic",
        "plainText",
        "number",
        "percent",
        "scientific",
        "currency",
        "date",
        "time",
        "increaseDecimal",
        "decreaseDecimal",
    ]


class CustomNumberFormat(TypedDict):
    type: Literal["custom"]
    pattern: str


NumberFormatMutation = Union[PresetNumberFormat, CustomNumberFormat]


class _CapturedFormat(TypedDict):
    rows: int
    columns: int
    formats: List[Dict[str, Any]]


class CapturedFormat(_CapturedFormat, total=False):
    """What :meth:`Workbook.capture_format` returns: ``rows`` x ``columns`` cell formats."""

    sourceStyles: List[Optional[int]]
    sourceFormats: List[Dict[str, Any]]


class SetCellOp(TypedDict):
    type: Literal["setCell"]
    sheet: int
    at: CellPoint
    cell: CellState


class InsertRowsOp(TypedDict):
    type: Literal["insertRows"]
    sheet: int
    at: int
    count: int


class DeleteRowsOp(TypedDict):
    type: Literal["deleteRows"]
    sheet: int
    at: int
    count: int


class InsertColsOp(TypedDict):
    type: Literal["insertCols"]
    sheet: int
    at: int
    count: int


class DeleteColsOp(TypedDict):
    type: Literal["deleteCols"]
    sheet: int
    at: int
    count: int


class _SetColWidthOp(TypedDict):
    type: Literal["setColWidth"]
    sheet: int
    col: int


class SetColWidthOp(_SetColWidthOp, total=False):
    """A missing or ``None`` width resets the column to the default."""

    width: Optional[float]


class SetColVisibilityOp(TypedDict):
    type: Literal["setColVisibility"]
    sheet: int
    col: int
    hidden: bool


class _SetRowHeightOp(TypedDict):
    type: Literal["setRowHeight"]
    sheet: int
    row: int


class SetRowHeightOp(_SetRowHeightOp, total=False):
    """A missing or ``None`` height resets the row to the default."""

    height: Optional[float]


class SetRowVisibilityOp(TypedDict):
    type: Literal["setRowVisibility"]
    sheet: int
    row: int
    hidden: bool


class _SetFreezePaneOp(TypedDict):
    type: Literal["setFreezePane"]
    sheet: int


class SetFreezePaneOp(_SetFreezePaneOp, total=False):
    """A missing or ``None`` pane unfreezes the sheet."""

    pane: Optional[FreezePane]


class SetHyperlinksOp(TypedDict):
    """Replaces every hyperlink on the sheet."""

    type: Literal["setHyperlinks"]
    sheet: int
    hyperlinks: List[Hyperlink]


class SetHyperlinkOp(TypedDict):
    type: Literal["setHyperlink"]
    sheet: int
    hyperlink: Hyperlink


class RemoveHyperlinkOp(TypedDict):
    type: Literal["removeHyperlink"]
    sheet: int
    range: CellRange


class SetAutoFilterOp(TypedDict):
    """``range`` is an A1 string, or ``None`` to remove the filter."""

    type: Literal["setAutoFilter"]
    sheet: int
    range: Optional[str]


class _AddConditionalFormatOp(TypedDict):
    type: Literal["addConditionalFormat"]
    sheet: int
    range: CellRange
    kind: Literal["cellIs", "expression"]
    formulas: List[str]
    format: DifferentialFormat


class AddConditionalFormatOp(_AddConditionalFormatOp, total=False):
    operator: str


class RemoveConditionalFormatOp(TypedDict):
    type: Literal["removeConditionalFormat"]
    sheet: int
    id: str


SetChartAnchorOp = TypedDict(
    "SetChartAnchorOp",
    {
        "type": Literal["setChartAnchor"],
        "sheet": int,
        "frame": str,
        "part": str,
        "from": ChartAnchor,
        "to": ChartAnchor,
    },
)


class MergeCellsOp(TypedDict):
    type: Literal["mergeCells"]
    sheet: int
    range: CellRange


class UnmergeCellsOp(TypedDict):
    type: Literal["unmergeCells"]
    sheet: int
    range: CellRange


class PatchRangeStyleOp(TypedDict):
    type: Literal["patchRangeStyle"]
    sheet: int
    range: CellRange
    patch: StylePatch


class SetRangeNumberFormatOp(TypedDict):
    type: Literal["setRangeNumberFormat"]
    sheet: int
    range: CellRange
    format: NumberFormatMutation


class ApplyRangeFormatOp(TypedDict):
    type: Literal["applyRangeFormat"]
    sheet: int
    range: CellRange
    format: CapturedFormat


class AddSheetOp(TypedDict):
    type: Literal["addSheet"]
    index: int
    name: str


class RemoveSheetOp(TypedDict):
    type: Literal["removeSheet"]
    index: int


MoveSheetOp = TypedDict(
    "MoveSheetOp", {"type": Literal["moveSheet"], "from": int, "to": int}
)


class RenameSheetOp(TypedDict):
    type: Literal["renameSheet"]
    sheet: int
    name: str


class _SetDefinedNameOp(TypedDict):
    type: Literal["setDefinedName"]
    name: str
    refers_to: str


class SetDefinedNameOp(_SetDefinedNameOp, total=False):
    """``scope`` is a sheet index for a sheet-local name, ``None`` for the workbook."""

    scope: Optional[int]
    hidden: Optional[bool]


class _DeleteDefinedNameOp(TypedDict):
    type: Literal["deleteDefinedName"]
    name: str


class DeleteDefinedNameOp(_DeleteDefinedNameOp, total=False):
    scope: Optional[int]


XlsxOp = Union[
    SetCellOp,
    InsertRowsOp,
    DeleteRowsOp,
    InsertColsOp,
    DeleteColsOp,
    SetColWidthOp,
    SetColVisibilityOp,
    SetRowHeightOp,
    SetRowVisibilityOp,
    SetFreezePaneOp,
    SetHyperlinksOp,
    SetHyperlinkOp,
    RemoveHyperlinkOp,
    SetAutoFilterOp,
    AddConditionalFormatOp,
    RemoveConditionalFormatOp,
    SetChartAnchorOp,
    MergeCellsOp,
    UnmergeCellsOp,
    PatchRangeStyleOp,
    SetRangeNumberFormatOp,
    ApplyRangeFormatOp,
    AddSheetOp,
    RemoveSheetOp,
    MoveSheetOp,
    RenameSheetOp,
    SetDefinedNameOp,
    DeleteDefinedNameOp,
]


class _WorkbookHyperlink(TypedDict):
    range: CellRange


class WorkbookHyperlink(_WorkbookHyperlink, total=False):
    """:meth:`Workbook.set_hyperlink`'s argument, camelCase like TypeScript's."""

    externalTarget: Optional[str]
    location: Optional[str]
    tooltip: Optional[str]
    display: Optional[str]


class ChartSeriesText(TypedDict):
    text: str


ChartSeriesRef = TypedDict("ChartSeriesRef", {"ref": str})
ChartSeriesName = Union[ChartSeriesText, ChartSeriesRef]


class SetChartTitle(TypedDict):
    sheet: int
    chart: int
    action: Literal["set_chart_title"]
    title: Optional[str]


class _SetChartSeries(TypedDict):
    sheet: int
    chart: int
    action: Literal["set_series"]
    series: int


class SetChartSeries(_SetChartSeries, total=False):
    name: ChartSeriesName
    categories: str
    values: str


ChartMutation = Union[SetChartTitle, SetChartSeries]
