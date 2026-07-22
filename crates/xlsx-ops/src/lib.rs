//! invertible op log: undo/redo, provenance-tagged proposals, and address
//! remapping. applying an op returns its inverse, so history is replay.

mod apply;
mod formatting;
mod input;
mod op;
mod proposals;
mod remap;
mod undo;

pub use apply::{InvertedOp, OpError, apply, apply_ops, remap_ref};
pub use formatting::{
    BorderLineStyle, BorderPatch, BorderPreset, CapturedFormat, HorizontalAlignment,
    NumberFormatMutation, StylePatch, StyleProperty, TextWrapping, VerticalAlignment,
};
pub use input::{ParsedInput, cell_state_for_input, cell_state_for_input_no_eval, parse_input};
pub use op::{CellState, Op, Provenance, Transaction};
pub use proposals::{Proposal, ProposalGhost, ProposalSet, ProposedEdit};
pub use undo::UndoStack;
