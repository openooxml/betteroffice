//! Bounded PresentationML parsing and byte-preserving package writes.

mod drawing;
mod error;
mod model;
mod package;
mod relationships;
mod theme;
mod xml;

pub use error::PptxError;
pub use model::*;
pub use package::{parse_pptx, parse_pptx_with_limits, write_pptx};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use xml::ParseLimits;
