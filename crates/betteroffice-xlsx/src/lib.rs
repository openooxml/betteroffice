//! Typed facade for opening, editing, calculating, rendering, and saving XLSX files.

mod error;
mod types;
mod workbook;

pub use error::Error;
pub use types::{
    CalculationOptions, CalculationResult, CellAddress, CellEdit, CellInput, MutationResult,
    ProposalAcceptance, ProposalEditInput, ProposalRequest, RenderOptions, RenderedPng, SheetInfo,
};
pub use workbook::{MAX_DISPLAY_CELLS, MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS, Workbook};

pub use xlsx_model::addr::AddrError;
pub use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ColId, DateSystem, ErrorValue, MAX_COLS, MAX_ROWS, RowId,
    Sheet, SheetId, Workbook as WorkbookModel,
};
pub use xlsx_ops::{CellState, Op, Proposal, ProposedEdit, Provenance, Transaction};
pub use xlsx_render::{
    Align, DisplayList, DrawCmd, GridGeometry, GridMeta, Rect, Viewport, viewport_for_range,
    viewport_for_used_range,
};

pub type Result<T> = std::result::Result<T, Error>;
