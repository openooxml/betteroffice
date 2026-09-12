//! Bounded Visio XML parsing and part-preserving package writes.

mod error;
mod model;
mod package;
mod relationships;
mod sheet;
mod xml;

pub use error::VsdxError;
pub use model::*;
pub use package::{parse_vsdx, parse_vsdx_with_limits, write_vsdx};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use sheet::*;
pub use xml::ParseLimits;
