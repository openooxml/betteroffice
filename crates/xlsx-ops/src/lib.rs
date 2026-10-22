//! invertible op log: undo/redo, provenance-tagged proposals, and address
//! remapping. applying an op returns its inverse, so history is replay.

mod apply;
mod conditional_formats;
mod formatting;
mod input;
mod op;
mod proposals;
mod remap;
mod undo;

pub use apply::{InvertedOp, OpError, apply, apply_in_place, apply_ops, remap_ref};
pub use conditional_formats::{
    ensure_conditional_formats_allow_edit, ensure_conditional_formats_allow_move,
};
pub use formatting::{
    BorderLineStyle, BorderPatch, BorderPreset, CapturedFormat, HorizontalAlignment, KeptFacets,
    NumberFormatMutation, PreservedCellFormat, StylePatch, StyleProperty, TextWrapping,
    VerticalAlignment,
};
pub use input::{
    ParsedInput, cell_state_for_input, cell_state_for_input_no_eval, cell_state_from_parsed,
    parse_input, parse_input_with_context,
};
pub use op::{CellState, Op, Provenance, Transaction};
pub use proposals::{Proposal, ProposalGhost, ProposalSet, ProposedEdit};
pub use remap::insertion_keeps_chart_anchor_on_grid;
pub use undo::UndoStack;
