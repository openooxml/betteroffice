//! Bounded Visio XML parsing and part-preserving package writes.

mod error;
mod model;
mod package;
mod patch;
mod relationships;
mod sheet;
mod xml;

pub use error::VsdxError;
pub use model::*;
pub use package::{
    CellLocator, CellRow, CellSheet, MutationGesture, SemanticCellEdit, StructuralEdit, parse_vsdx,
    parse_vsdx_with_limits, save_semantic_cell_edits, save_structural_edits, validate_structure,
    write_vsdx,
};
pub use patch::{
    CellAttribute, CellEdit, ElementSpan, MAX_PATCH_BYTES, MAX_PATCH_EDITS, SourceSpan, SpanEdit,
    apply_span_edits,
};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use sheet::*;
pub use xml::ParseLimits;
