use xlsx_model::{CellRange, CellRef, SheetId};

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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProposalEditInput {
    pub sheet: SheetId,
    pub cell: CellRef,
    pub input: String,
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
