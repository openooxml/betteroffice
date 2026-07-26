//! The `drawingml/2006/chart` part: model, parsing, and plot geometry.
//!
//! `c:chartSpace` is schema-identical across docx, xlsx and pptx, so nothing
//! here is format-specific. Hosts supply XML access via [`ChartXml`] and
//! translate the [`PlotOp`] list into their own primitives.

mod geometry;
mod model;
mod parse;

pub use geometry::*;
pub use model::*;
pub use parse::*;
