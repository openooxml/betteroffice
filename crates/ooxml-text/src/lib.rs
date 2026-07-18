//! Pure-Rust text shaping and measurement for the DOCX layout engine.
//!
//! This crate is the `ooxml-text` module described in
//! `openspec/changes/rust-canvas-engine/design.md`. It owns everything the
//! layout engine needs to turn run text into positioned glyphs and line-break
//! decisions, with no browser APIs in the loop:
//!
//! - [`FontStore`] — registry over raw font **bytes** (never font names) with
//!   per-font metrics (`head`/`hhea`/`OS/2`), cmap lookup, advance widths, and
//!   an ordered fallback-chain resolver.
//! - [`shape`] / [`shape_with_direction`] — OpenType shaping via rustybuzz,
//!   returning cluster-mapped glyphs with advances/offsets scaled to the
//!   requested size.
//! - [`break_opportunities`] — UAX-14 line-break opportunities via
//!   `unicode-linebreak` (surrogate-safe, CJK-aware).
//! - [`bidi_paragraphs`] — paragraph-level Unicode Bidirectional Algorithm
//!   runs via `unicode-bidi`.
//! - [`word_metrics`] — the Word-specific measurement rules: single-spacing
//!   line boxes from OS/2 win metrics ([`single_line_box`]), auto/exact/
//!   atLeast line rules ([`apply_spacing_rule`]), justification gating and
//!   space-stretch ([`line_is_justified`], [`stretch_spaces`]), the w:kern
//!   threshold ([`kern_enabled`], [`kern_features`]), and the settings.xml
//!   compat flags that feed them ([`CompatFlags`]). Snap-to-grid (w:docGrid)
//!   remains a documented TODO there.
//! - [`outline`] — glyph outline extraction ([`FontStore::outline_glyph`]):
//!   font-unit path commands ([`PathCmd`]) from the same skrifa bytes the
//!   metrics came from, for the canvas renderer's `Path2D` glyph pipeline.
//!
//! Design constraint (load-bearing): callers hand this crate font *bytes* plus
//! a fallback chain of [`FontId`]s. Resolving a `w:rFonts` name to bytes —
//! embedded `.odttf`, bundled metric-compatible fonts, Local Font Access, or
//! browser-measured fallback — happens entirely on the host side. That keeps
//! this crate deterministic and identical across web and native shells.
//!
//! No `wasm-bindgen` here by design: this mirrors `docx-layout`'s
//! pure-`layout_to_json` lesson. A thin WASM facade crate can wrap this later.

#![allow(clippy::type_complexity)]

pub mod bidi;
pub mod font_store;
pub mod line_break;
pub mod measure;
pub mod outline;
pub mod shape;
pub mod word_metrics;

pub use bidi::{
    BaseDirection, BidiParagraph, BidiRun, bidi_paragraphs, level_is_rtl, visual_order_for_levels,
};
pub use font_store::{FontError, FontId, FontMetrics, FontStore};
pub use line_break::{BreakOpportunity, break_opportunities};
pub use measure::{
    MeasureError, MeasureInput, ParagraphExtentOut, TypesetRowOut, measure_paragraph,
    measure_paragraph_json,
};
pub use outline::{GlyphOutline, PathCmd};
pub use shape::{ShapeDirection, ShapeFeature, ShapedGlyph, shape, shape_with_direction};
pub use word_metrics::{
    CompatFlags, LineBox, LineSpacingRule, apply_spacing_rule, kern_enabled, kern_features,
    line_is_justified, single_line_box, stretch_spaces,
};
