use serde::Serialize;
use xlsx_model::{CellRange, CellRef, SheetId};

use xlsx_ops::{
    BorderLineStyle, BorderPreset, HorizontalAlignment, NumberFormatMutation, TextWrapping,
    VerticalAlignment,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UpdateOrigin {
    Local,
    Remote,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UpdateEvent {
    pub update: Vec<u8>,
    pub origin: UpdateOrigin,
}

#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct CalculationOptions {
    pub now_serial: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddress {
    pub sheet: SheetId,
    pub cell: CellRef,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellInput {
    pub cell: CellRef,
    pub input: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellEdit {
    pub cell: CellRef,
    pub input: String,
    pub is_formula: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SheetInfo {
    pub sheet_names: Vec<String>,
    pub active_sheet: SheetId,
    pub content_width: f32,
    pub content_height: f32,
    pub frozen_rows: u32,
    pub frozen_cols: u32,
    pub initial_scroll_x: f32,
    pub initial_scroll_y: f32,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CalculationResult {
    pub changed: Vec<CellAddress>,
    pub cycle_cells: Vec<CellAddress>,
    pub limited_cells: Vec<CellAddress>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MutationResult {
    pub applied: bool,
    pub changed: Vec<CellAddress>,
    pub cycle_cells: Vec<CellAddress>,
    pub limited_cells: Vec<CellAddress>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum NumberFormatKind {
    Automatic,
    PlainText,
    Number,
    Percent,
    Scientific,
    Currency,
    Date,
    Time,
    Custom,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SelectionFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<NumberFormatKind>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format_pattern: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_preset: Option<BorderPreset>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_style: Option<BorderLineStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_alignment: Option<HorizontalAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<VerticalAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_wrapping: Option<TextWrapping>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HistoryState {
    pub can_undo: bool,
    pub can_redo: bool,
    pub undo_depth: usize,
    pub redo_depth: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProposalEditInput {
    pub sheet: SheetId,
    pub cell: CellRef,
    pub input: String,
    pub number_format: Option<NumberFormatMutation>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProposalRequest {
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposalEditInput>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProposalAcceptance {
    pub proposal_id: String,
    pub mutation: MutationResult,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RenderOptions {
    pub range: Option<CellRange>,
    pub scale: f32,
    pub max_width: Option<u32>,
    pub max_height: Option<u32>,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            range: None,
            scale: 1.0,
            max_width: None,
            max_height: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RenderedPng {
    pub bytes: Vec<u8>,
    pub width: u32,
    pub height: u32,
}
