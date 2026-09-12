//! Resolved, non-mutating views over `vsdx_parse` sheets.

mod geometry;
mod inheritance;
mod model;
mod text;

#[cfg(test)]
mod tests;

pub use geometry::*;
pub use inheritance::*;
pub use model::*;
