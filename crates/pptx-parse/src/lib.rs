//! Bounded PresentationML parsing and part-preserving package writes.

mod chart;
mod drawing;
mod error;
mod model;
mod package;
mod relationships;
mod theme;
mod write;
mod xml;

pub use error::PptxError;
pub use model::*;
pub use package::{parse_pptx, parse_pptx_with_limits, write_pptx};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use write::{
    DeckWrite, InheritedTransform, ParagraphWrite, RunWrite, ShapeAdd, ShapePatch, ShapeWrite,
    SlideWrite, TextTarget, TextWrite, write_pptx_with_edits,
};
pub use xml::ParseLimits;
