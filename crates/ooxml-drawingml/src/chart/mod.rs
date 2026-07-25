//! The `drawingml/2006/chart` part: model, parsing, and plot geometry.
//!
//! `c:chartSpace` is schema-identical across docx, xlsx and pptx, so nothing
//! here is format-specific.

mod model;

pub use model::*;
