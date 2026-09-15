//! Resolved, non-mutating views over `vsdx_parse` sheets.

mod connectivity;
mod containers;
mod geometry;
mod inheritance;
mod model;
mod shape_data;
mod text;

#[cfg(test)]
mod tests;

pub use connectivity::*;
pub use containers::*;
pub use geometry::*;
pub use inheritance::*;
pub use model::*;
pub use shape_data::*;
