//! Typed facade for opening, editing, calculating, rendering, and saving XLSX files.

mod authority;
mod error;
mod sheet_json;
mod structured;
mod types;
mod workbook;

pub use error::Error;
pub use structured::{
    DEFAULT_EXPORT_MAX_BYTES, DEFAULT_EXPORT_MAX_CELLS, DEFAULT_MARKDOWN_MAX_CELLS,
    DEFAULT_MARKDOWN_MAX_COLUMNS, DEFAULT_MARKDOWN_MAX_ROWS, MAX_EXPORT_BYTES, MAX_EXPORT_CELLS,
    MAX_MARKDOWN_CELLS, MAX_MARKDOWN_ROWS, MIN_MARKDOWN_BYTES, XlsxAnchor, XlsxAnchorScope,
    XlsxCalculationFreshness, XlsxCalculationPolicy, XlsxCalculationState, XlsxCellMerge,
    XlsxDateSystem, XlsxExport, XlsxExportCell, XlsxExportDefinedName, XlsxExportDiagnostic,
    XlsxExportDiagnosticCode, XlsxExportFailure, XlsxExportFailureCode, XlsxExportHyperlink,
    XlsxExportIncluded, XlsxExportMerge, XlsxExportObject, XlsxExportOptions, XlsxExportRefusal,
    XlsxExportResult, XlsxExportScope, XlsxExportSeverity, XlsxExportSheet, XlsxExportTable,
    XlsxExportValue, XlsxFormulaResult, XlsxMarkdownAnchor, XlsxMarkdownContent,
    XlsxMarkdownOptions, XlsxObjectKind, XlsxSchemaVersion, XlsxSheetIdentity, XlsxSheetKind,
    XlsxSheetVisibility, XlsxSourcePart, XlsxStructuredContent, export_result_json,
    export_xlsx_markdown, export_xlsx_markdown_json, export_xlsx_structured,
    export_xlsx_structured_json, render_xlsx_markdown, render_xlsx_markdown_json,
};
pub use types::{
    CalculationOptions, CalculationResult, CellAddress, CellEdit, CellInput, EditProfile,
    EditStage, HistoryState, MutationResult, NumberFormatKind, ProposalAcceptance,
    ProposalEditInput, ProposalRequest, RenderOptions, RenderedPng, SelectionFormatting, SheetInfo,
    TextSearchMatch, UpdateEvent, UpdateOrigin,
};
pub use workbook::batch::{
    CalculationRequest, CellGuard, DocumentVersion, EditApplication, EditCalculation, EditFailure,
    EditFailureCode, EditHistory, EditOperation, EditOutcome, EditPreview, EditReceipt,
    EditRefusal, EditRequest, EditSource, EditStep, EditValidation, MAX_REQUEST_BYTES,
    MAX_RESPONSE_BYTES, StepGuard, ValidationOutcome, outcome_json,
};
pub use workbook::target::{
    CellPosition, CellRead, CellTarget, CellsRead, FindMatch, FindOutcome, FindRequest,
    RangeAddress, RangeRead, RangeTarget, ReadCalculation, ReadOutcome, ReadRequest, SheetEntry,
    TextFound,
};
pub use workbook::{
    DEFAULT_TEXT_SEARCH_LIMIT, MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID,
    MAX_COLLABORATION_STATE_VECTOR_ENTRIES, MAX_DISPLAY_CELLS, MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS,
    UpdateSubscription, Workbook,
};

pub use xlsx_model::addr::AddrError;
pub use xlsx_model::{
    AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, Cell, CellRange, CellRef, CellValue,
    ChartAnchor, ChartRef, ChartRefKind, ColId, ColStyle, DateSystem, DefinedName, ErrorValue,
    FreezePane, Hyperlink, MAX_COLS, MAX_ROWS, RowId, Sheet, SheetChart, SheetFormat, SheetId,
    Stylesheet, Workbook as WorkbookModel,
};
pub use xlsx_ops::{
    BorderLineStyle, BorderPatch, BorderPreset, CapturedFormat, CellState, HorizontalAlignment,
    NumberFormatMutation, Op, Proposal, ProposedEdit, Provenance, StylePatch, StyleProperty,
    TextWrapping, Transaction, VerticalAlignment,
};
pub use xlsx_render::{
    Align, ChartA11yAttrs, ChartRegion, DisplayList, DrawCmd, GridGeometry, GridMeta,
    HyperlinkRegion, PathStroke, PrintMetrics, Rect, RenderError, Viewport, viewport_for_range,
    viewport_for_used_range, viewport_for_used_range_within,
};

pub type Result<T> = std::result::Result<T, Error>;
