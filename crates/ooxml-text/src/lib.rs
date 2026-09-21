//! Pure-Rust text shaping and measurement for the DOCX layout engine.
//!
//! It owns everything the layout engine needs to turn run text into
//! positioned glyphs and line-break
//! decisions, with no browser APIs in the loop:
//!
//! - [`FontStore`] — registry over raw font **bytes** (never font names) with
//!   per-font metrics (`head`/`hhea`/`OS/2`), cmap lookup, advance widths, and
//!   an ordered fallback-chain resolver.
//! - [`shape()`] / [`shape_with_direction`] — OpenType shaping via rustybuzz,
//!   returning cluster-mapped glyphs with advances/offsets scaled to the
//!   requested size.
//! - [`break_opportunities`] — UAX-14 line-break opportunities via
//!   `unicode-linebreak` (surrogate-safe, CJK-aware).
//! - [`bidi_paragraphs`] — paragraph-level Unicode Bidirectional Algorithm
//!   runs via `unicode-bidi`.
//! - [`word_metrics`] — the Word-specific measurement rules: single-spacing
//!   line boxes from OS/2 win metrics ([`single_line_box`], with disabled
//!   metric experiments behind [`CompatFlags`]), auto/exact/atLeast line
//!   rules ([`apply_spacing_rule`]), justification gating and
//!   space-stretch ([`line_is_justified`], [`stretch_spaces`]), the w:kern
//!   threshold ([`kern_enabled`], [`kern_features`]), document-grid line
//!   snapping ([`snap_line_box`], [`snap_line_height`], active only for an
//!   activating `w:docGrid` type with per-paragraph/per-run `w:snapToGrid`
//!   opt-outs), and the settings.xml
//!   compat flags that feed them ([`CompatFlags`]). Snap-to-grid (w:docGrid)
//!   rounds the *content* box up to a whole number of grid rows via
//!   [`snap_line_box`] when the caller supplies an activating grid pitch, so
//!   the `auto` multiple then scales the quantized pitch.
//! - [`caps`] — the casing `w:caps`/`w:smallCaps` and `a:rPr/@cap` share:
//!   language-aware uppercasing ([`uppercase_for_language`]) and the
//!   synthesized small-cap advance scales.
//! - [`auto_space`] — `w:autoSpaceDE`/`w:autoSpaceDN`: the quarter-em Word
//!   inserts where East Asian text meets Latin letters or digits.
//! - [`word_fonts`] — the vertical metrics of the East Asian faces Word
//!   ships, so a substituted face measures as the one the document named.
//! - [`symbol_font`] — what a Wingdings or Webdings character actually
//!   addresses, and the nearest covered Unicode character to draw for it.
//! - [`outline`] — glyph outline extraction ([`FontStore::outline_glyph`]):
//!   font-unit path commands ([`PathCmd`]) from the same skrifa bytes the
//!   metrics came from, for the canvas renderer's `Path2D` glyph pipeline.
//!
//! Design constraint (load-bearing): callers hand this crate font *bytes* plus
//! a fallback chain of [`FontId`]s. Resolving a `w:rFonts` name to bytes —
//! embedded `.odttf`, bundled metric-compatible fonts, Local Font Access, or
//! browser-measured fallback — happens entirely on the host side. That keeps
//! this crate deterministic and identical across web and native shells.

#![allow(clippy::type_complexity)]

pub mod auto_space;
pub mod bidi;
pub mod caps;
pub mod font_store;
pub mod line_break;
pub mod measure;
pub mod outline;
pub mod shape;
pub mod symbol_font;
pub mod word_fonts;
pub mod word_metrics;

pub use auto_space::{AUTO_SPACE_EM, AutoSpace};
pub use bidi::{
    BaseDirection, BidiParagraph, BidiRun, bidi_paragraphs, level_is_rtl, visual_order_for_levels,
};
pub use caps::{
    BROWSER_SMALL_CAPS_ADVANCE_SCALE, WORD_SMALL_CAPS_ADVANCE_SCALE, uppercase_for_language,
};
pub use font_store::{FontError, FontId, FontMetrics, FontStore, RequestedLineMetrics};
pub use line_break::{BreakOpportunity, break_opportunities};
pub use measure::{
    FontChains, MeasureError, MeasureInput, MeasureRequest, ParagraphExtentOut, TypesetRowOut,
    measure_paragraph, measure_paragraph_json, measure_paragraph_typed,
};
pub use outline::{GlyphOutline, PathCmd};
pub use shape::{ShapeDirection, ShapeFeature, ShapedGlyph, shape, shape_with_direction};
pub use symbol_font::SymbolFont;
pub use word_metrics::{
    CompatFlags, LineBox, LineSpacingRule, apply_spacing_rule, kern_enabled, kern_features,
    line_is_justified, single_line_box, snap_line_box, snap_line_height, stretch_spaces,
};
