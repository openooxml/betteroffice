//! Resolved, non-mutating views over `vsdx_parse` sheets.

mod connectivity;
mod geometry;
mod inheritance;
mod model;
mod text;

#[cfg(test)]
mod tests;

pub use connectivity::*;
pub use geometry::*;
pub use inheritance::*;
pub use model::*;
