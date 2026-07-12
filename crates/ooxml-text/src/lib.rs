//! Shared OOXML text shaping and measurement.

pub mod bidi;
pub mod font_store;
pub mod line_break;
pub mod line_metrics;
pub mod measure;
pub mod outline;
pub mod shape;

pub use bidi::{
    BaseDirection, BidiParagraph, BidiRun, bidi_paragraphs, level_is_rtl, visual_order_for_levels,
};
pub use font_store::{FontError, FontId, FontMetrics, FontStore};
pub use line_break::{BreakOpportunity, break_opportunities};
pub use line_metrics::{
    CompatFlags, LineBox, LineSpacingRule, apply_spacing_rule, kern_enabled, kern_features,
    line_is_justified, single_line_box, stretch_spaces,
};
pub use measure::{
    MeasureError, MeasureInput, ParagraphExtentOut, TypesetRowOut, measure_paragraph,
    measure_paragraph_json,
};
pub use outline::{GlyphOutline, PathCmd};
pub use shape::{ShapeDirection, ShapeFeature, ShapedGlyph, shape, shape_with_direction};
