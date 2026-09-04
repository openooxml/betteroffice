//! Display-list builder: `{ measured, options } + Layout -> DisplayList`.
//!
//! Takes the same measured/options envelope pagination consumes plus the
//! finished [`crate::types::Layout`], and compiles one [`DisplayPage`] of
//! renderer-agnostic paint primitives per laid-out page. A display list is
//! always derived from a `Layout`; nothing here decides where content goes.
//!
//! Every text-bearing primitive carries `docStart` / `docEnd` and its block
//! identity. That is what hit testing, selection and the accessibility mirror
//! resolve against, so a primitive without those is invisible to interaction.
//!
//! # Paint order
//!
//! A page's `primitives` are emitted back to front: watermark, behind-document
//! floating images, the layout's fragments in order, in-front floating images,
//! then column separators. Inside a fragment the order is shading, borders,
//! then line content.
//!
//! `background`, `page_borders`, `header`, `footer` and `note_areas` are
//! separate fields rather than entries in that stream, so the consumer places
//! them. Page borders are not one layer: each carries a `z_order` of `back` or
//! `front` taken from `w:pgBorders/@zOrder` (anything but `back` is `front`),
//! so a border may sit behind page content or over it.
//!
//! # Text
//!
//! Run positions come from resolved line segments, and decorations — underline,
//! strike, highlight, comment range, revision washes — ride with the run they
//! belong to. A run becomes a [`GlyphRunPrimitive`] when the input supplies
//! font chains and the run's chain resolves against the measurement font store;
//! any miss (no fonts, no chain, empty text, a shaping failure) falls back to a
//! [`TextRunPrimitive`] carrying a CSS font shorthand. A glyph run is
//! single-font by contract, so a run spanning fallback fonts splits into one
//! glyph run per maximal same-font subrange.
//!
//! Rows measured with authoritative run, cluster and bidi metadata use it
//! directly. Rows without it fall back to a deterministic proportional split:
//! fixed-width items (tabs, inline images) claim their width first, and the
//! line's remaining measured width is distributed across the text-ish segments
//! by character count.
//!
//! PAGE and NUMPAGES fields whose per-page resolved widths the input supplies
//! leave that pool and become fixed-width at the page's own width, with the
//! fallback width they contributed to the measure subtracted from the pool
//! baseline — which is what lets a centred or right-aligned field line
//! re-centre per page instead of tracking the once-measured fallback. DATE and
//! TIME fields always render their stored fallback text rather than the current
//! clock, so a display list is reproducible. A field's instruction is carried
//! for announcement only and is never evaluated.
//!
//! Justified lines (`jc` of `both` or `distribute`) spread the line's slack
//! evenly across U+0020 clusters. Glyph positions already fold that stretch in,
//! so the emitted `wordSpacing` is an interchange hint, not something a
//! renderer re-applies. On the first line the hanging or first-line indent
//! shifts the text; when a visible list marker is present it occupies the hang,
//! and body text starts at the text indent instead of being pulled left.
//!
//! # Tables
//!
//! A table fragment paints the row window its page shows, stacking a repeated
//! header band above it. Vertically merged cells are re-emitted as continuation
//! slices on later pages: those are re-paints, so they carry no document
//! positions and are not selectable. Repeated header rows are literal re-paints
//! and are deliberately *not* flagged as continuations. Shared cell edges
//! collapse to a single border, and a fragment cut by a page break is closed
//! with its own cut edges.
//!
//! Inside a cell, paragraph spacing collapses as it does in body flow, drawn
//! border widths and padding inset the content box, and `w:vAlign` offsets the
//! leftover slack — except that content which fills or overflows the box stays
//! top-aligned, matching where the paginator broke it. Text boxes nested inside
//! table cells emit nothing.
//!
//! # Numbers
//!
//! Coordinates round to three decimals, and integral floats are canonicalized
//! to integers so the fields typed as integers survive a round trip through a
//! host with a single number type.

use ooxml_drawingml::GeometryPathCommand;
use ooxml_drawingml::chart::{
    PlotAxis, PlotAxisKind, PlotAxisRange, PlotAxisTitles, PlotChart, PlotChartText,
    PlotDataLabels, PlotGroup, PlotLegend, PlotMarker, PlotMarkerSymbol, PlotOp, PlotPoint,
    PlotRect, PlotSeries, PlotSink, PlotTextStyle, chart_aria_label, plot_chart_into,
};
use serde::de::DeserializeOwned;
use serde::{Deserialize, Serialize};
use serde_json::{Number, Value};
use std::collections::{HashMap, HashSet};

/// A compiled document: one entry per laid-out page, in document order.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DisplayList {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contract_version: Option<u32>,
    pub pages: Vec<DisplayPage>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DisplayPage {
    pub page_index: u64,
    pub width: Number,
    pub height: Number,
    /// Authored body content box and column boxes. These are explicit display
    /// metadata because interaction code must not infer Word margins from the
    /// minimum coordinates of whichever primitives happen to paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_bounds: Option<DisplayBounds>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub column_bounds: Vec<DisplayBounds>,
    /// Effective section and page-number state from pagination.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_page_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_page_number: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_label: Option<String>,
    /// paint order
    pub primitives: Vec<Primitive>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_borders: Vec<PageBorderPrimitive>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header: Option<HfRegion>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer: Option<HfRegion>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub note_areas: Vec<NoteRegion>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DisplayBounds {
    pub x: Number,
    pub y: Number,
    pub width: Number,
    pub height: Number,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct NoteRegion {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<u64>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub separator_primitives: Vec<Primitive>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub primitives: Vec<Primitive>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub note_ids: Vec<i64>,
    /// Per-note backlink metadata: the body-document anchor range and the
    /// formatted label, so a note can be linked back to its reference mark.
    /// Emitted only for notes that carry anchor or label data.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub notes: Vec<NoteRegionNote>,
}

/// Backlink metadata for one note region entry.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct NoteRegionNote {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<i64>,
    /// body-doc range of the reference mark anchoring this note
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor_doc_start: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor_doc_end: Option<i64>,
    /// formatted reference label (display number / custom mark); file-derived
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<String>,
}

/// Header or footer band in page coordinates.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct HfRegion {
    pub r_id: String,
    pub kind: HfKind,
    pub y: Number,
    pub height: Number,
    /// paint order
    pub primitives: Vec<Primitive>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum HfKind {
    #[serde(rename = "header")]
    Header,
    #[serde(rename = "footer")]
    Footer,
}

/// One paint operation. A renderer walks these in order and needs no knowledge
/// of DOCX to draw them.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(tag = "kind")]
pub enum Primitive {
    #[serde(rename = "text")]
    Text(TextRunPrimitive),
    #[serde(rename = "glyphRun")]
    GlyphRun(GlyphRunPrimitive),
    #[serde(rename = "rect")]
    Rect(RectPrimitive),
    #[serde(rename = "line")]
    Line(LinePrimitive),
    #[serde(rename = "image")]
    Image(ImagePrimitive),
    #[serde(rename = "shape")]
    Shape(ShapePrimitive),
    #[serde(rename = "decoration")]
    Decoration(DecorationPrimitive),
}

/// Everything a primitive carries about the document content it came from:
/// its position range, block and paragraph identity, enclosing table cell,
/// content-control ancestry, revision and comment state, and the per-class
/// paint metadata flattened alongside. This is the only channel through which
/// a painted primitive can be mapped back to the document.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocAttrs {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub block_id: Option<Number>,
    /// raw string block id, emitted when the id is NOT numeric (the live
    /// pipeline's compound `block-N` keys). Consumers group by
    /// `blockKey ?? String(blockId)` — exactly one of the two is present
    /// whenever the primitive has block identity.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub block_key: Option<String>,
    /// Exact document range of the owning paragraph fragment. A primitive's own
    /// range describes inline content and normally starts at the paragraph's
    /// `pmStart + 1`, whereas a paragraph wrapper starts at the fragment's
    /// `pmStart`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fragment_doc_start: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fragment_doc_end: Option<i64>,
    /// Stable Word `w14:paraId` of the enclosing paragraph.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    /// Measured line window `[from_line, to_line)` of the paragraph fragment
    /// this primitive belongs to. Stamped only where a fragment surfaces as a
    /// paragraph wrapper — body, header/footer and text-box content. Table-cell
    /// paragraphs are omitted, because they surface as ARIA cells and have no
    /// fragment of their own to carry the window.
    /// Omitted when no stamped range exists.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from_line: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to_line: Option<u64>,
    /// Resolved line index within the owning paragraph.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_index: Option<u64>,
    /// table cell the primitive paints inside (0-based grid coordinates)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell: Option<TableCellRef>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment_ids: Option<Vec<String>>,
    /// inert field identity when this primitive paints a field result — the
    /// a11y mirror announces it; the instruction is NEVER parsed/executed.
    /// Additive + serde-optional: field-free fixtures stay byte-identical.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field: Option<FieldMetadata>,
    /// footnote/endnote reference identity when this primitive is the body
    /// reference mark (note backlinks). Additive + serde-optional.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_ref: Option<NoteRefMetadata>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<Revision>,
    /// Synthetic numbering glyph emitted before the first line of a list
    /// paragraph. The mirror uses this to expose the stable list-marker class.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_marker: Option<bool>,
    /// Pending tracked-change paint for the synthetic numbering glyph.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_marker_revision: Option<RevisionKind>,
    /// structural tracked-change cue on a paragraph mark or table structure
    /// (paragraph mark, whole table, row, or cell). Additive + optional so
    /// display-list snapshots that carry only run-level revisions stay
    /// byte-identical.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub structural_revision: Option<StructuralRevision>,
    /// sanitized hyperlink target for clickable text/image primitives.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub href: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_target: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_history: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_doc_location: Option<String>,
    /// innermost block-level content-control identity.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt: Option<SdtAttrs>,
    /// Full outer-to-inner content-control ancestry.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sdt_path: Vec<SdtAttrs>,
    /// inline content-control widget metadata when this text primitive is its glyph.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inline_sdt_widget: Option<InlineSdtWidgetAttrs>,
    /// accessibility summary for primitives that compose one chart block.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart: Option<ChartA11yAttrs>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub aria_label: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub aria_description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden_object: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<CommentMetadata>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clip_group: Option<ClipGroupMetadata>,
    /// Leader glyph metadata shared by text and glyph primitives.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_glyphs: Option<LeaderGlyphMetadata>,
    /// Optional decoration metadata.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight_slice: Option<HighlightSliceMetadata>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<DisplayBorderStyle>,
    /// Fields belonging to one primitive class are flattened through the shared
    /// attrs, so a constructor that fills only the common fields stays valid.
    #[serde(default, skip_serializing_if = "Option::is_none", rename = "opacity")]
    pub primitive_opacity: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none", rename = "flipH")]
    pub image_flip_h: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none", rename = "flipV")]
    pub image_flip_v: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_frame: Option<ContentFrame>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effects: Vec<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint: Option<Value>,
    /// Lossless DrawingML stroke details beyond the plain colour/width/dash
    /// triple: compound, alignment, caps, joins, arrows and custom dashes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_paint: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_extent: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_scene: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_body_properties: Option<Value>,
    /// GlyphRun-only member flattened through the shared attrs (same pattern
    /// as the image/shape members above): the resolved CSS font shorthand the
    /// canvas fillText safety net uses when glyph outlines are unavailable,
    /// so the fallback keeps the measured face instead of generic sans-serif.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fallback_font: Option<String>,
    /// Text/GlyphRun-only member: modern w14 text effects payload
    /// (glow/shadow/reflection/textFill/textOutline), passed through losslessly
    /// from `RunFormatting.modernEffects`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modern_effects: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table: Option<TableMetadata>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableMetadata {
    pub table_id: String,
    pub row_start: u64,
    pub row_end: u64,
    pub row_count: u64,
    pub column_count: u64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_count: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_table_id: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct CommentMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub palette_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selected: Option<bool>,
    /// reviewer display name (file-derived, attacker-controlled)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author_name: Option<String>,
    /// comment date string, passed through verbatim
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    /// plain-text comment body excerpt, capped at [`MAX_COMMENT_TEXT_CHARS`]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// total replies in the thread (may exceed `replies.len()` when capped)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reply_count: Option<u64>,
    /// reply summaries in thread order, capped at [`MAX_COMMENT_REPLIES`]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub replies: Vec<CommentReplyMetadata>,
}

/// One reply summary inside [`CommentMetadata`].
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct CommentReplyMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
}

/// Inert field identity carried by a primitive.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct FieldMetadata {
    /// resolved category: PAGE, NUMPAGES, DATE, TIME or OTHER
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
    /// raw Word field type token (e.g. TOC, PAGEREF, REF)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<String>,
    /// raw instruction text, capped at [`MAX_FIELD_INSTRUCTION_CHARS`]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub instruction: Option<String>,
}

/// Footnote or endnote identity on a body reference mark.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct NoteRefMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<i64>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ClipGroupMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clip: Option<ClipRect>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<Number>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct ClipRect {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<Number>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct LeaderGlyphMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glyph: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline_y: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub advance: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct HighlightSliceMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_start: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_end: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascent: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub descent: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub includes_trailing_whitespace: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct ChartA11yAttrs {
    pub label: String,
}

/// The grid position of the cell a primitive paints inside.
///
/// `row` and `col` are 0-based anchor-grid coordinates with vertical merges and
/// column spans already resolved, so a merged cell keeps its anchor row, and
/// spans are always at least 1. `continuation` marks the synthetic slice of a
/// vertically merged cell re-painted on a later page: not selectable, and with
/// document positions stripped. A repeated header row is a literal re-paint and
/// is deliberately not flagged.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct TableCellRef {
    pub row: u64,
    pub col: u64,
    pub row_span: u64,
    pub col_span: u64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub continuation: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub is_header: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub repeated_header: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<bool>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header_ids: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub owns_top_border: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub owns_right_border: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub owns_bottom_border: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub owns_left_border: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SdtAttrs {
    pub group_id: String,
    pub sdt_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub depth: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lock: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bound: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub repeating_item: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct InlineSdtWidgetAttrs {
    pub kind: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_kind: Option<String>,
    pub group_id: String,
    pub pos: i64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_id: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selected_index: Option<i64>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub list_items: Vec<InlineSdtListItemAttrs>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locked: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct InlineSdtListItemAttrs {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Revision {
    pub author: String,
    pub date: String,
    pub revision_id: String,
    pub kind: RevisionKind,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq)]
pub enum RevisionKind {
    #[serde(rename = "ins")]
    Ins,
    #[serde(rename = "del")]
    Del,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct StructuralRevision {
    pub scope: StructuralRevisionScope,
    pub author: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub revision_id: String,
    pub kind: StructuralRevisionKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_index: Option<u64>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuralRevisionScope {
    #[serde(rename = "pmark")]
    Pmark,
    #[serde(rename = "table")]
    Table,
    #[serde(rename = "row")]
    Row,
    #[serde(rename = "cell")]
    Cell,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuralRevisionKind {
    #[serde(rename = "ins")]
    Ins,
    #[serde(rename = "del")]
    Del,
    #[serde(rename = "merge")]
    Merge,
}

/// A run of text painted from a font shorthand and a measured advance, used
/// wherever shaping is unavailable for the run.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct TextRunPrimitive {
    pub text: String,
    /// pen origin
    pub x: Number,
    pub baseline_y: Number,
    /// measured advance of the whole run
    pub width: Number,
    /// Guessed horizontal paint slot. Missing y/h leaves ink vertically unbounded.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paint_clip: Option<ClipRect>,
    /// CSS font shorthand.
    pub font: String,
    pub color: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub letter_spacing: Option<Number>,
    /// Extra advance added after each U+0020 cluster (px), set only on
    /// justified lines (`jc` of `both` or `distribute`). The primitive's
    /// `width` already includes the same stretch, so this is an interchange
    /// hint. Absent means no stretch.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_spacing: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation_deg: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_scale: Option<Number>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub all_caps: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub small_caps: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_shadow: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub text_outline: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_effect: Option<String>,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

/// A shaped run of positioned glyphs, single-font by contract.
///
/// A renderer paints each glyph as an outline from `font_id`'s bytes, while
/// `text` keeps the real characters for accessibility and copying. Emitted only
/// when the measurement font store is populated and the run's font chain
/// resolves; otherwise the builder emits a [`TextRunPrimitive`]. A run spanning
/// several fallback fonts becomes one glyph run per maximal same-font subrange.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct GlyphRunPrimitive {
    /// id in the measurement `FontStore` (the same registry `fontChains`
    /// references).
    pub font_id: u32,
    /// font size in px.
    pub size: f64,
    pub color: String,
    /// SOURCE TEXT — the a11y mirror renders this as real characters and the
    /// glyph `cluster`s index into it. REQUIRED.
    pub text: String,
    pub glyphs: Vec<PlacedGlyph>,
    /// Guessed horizontal paint slot. Missing y/h leaves ink vertically unbounded.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paint_clip: Option<ClipRect>,
    /// extra advance added after each U+0020 cluster on a justified line (px) —
    /// equivalent to [`TextRunPrimitive::word_spacing`]; the glyph `x` positions
    /// already fold this stretch in, so it is an interchange hint, not
    /// re-applied by the renderer.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_spacing: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation_deg: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_scale: Option<Number>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub all_caps: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub small_caps: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_shadow: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub text_outline: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_effect: Option<String>,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

/// One positioned glyph inside a [`GlyphRunPrimitive`].
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PlacedGlyph {
    /// glyph id within the run's `font_id`.
    pub id: u32,
    /// pen origin x, page-local px (accumulated advances + x_offset + any
    /// justification stretch).
    pub x: f64,
    /// baseline y, page-local px (includes the shaped y_offset).
    pub y: f64,
    /// BYTE index into the run's `text` of the cluster this glyph belongs to
    /// (the glyph↔char map for hit-testing / a11y).
    pub cluster: u32,
    /// pen advance for this glyph in page-local px — how far the pen moves after
    /// painting it (the shaped `x_advance` plus any justification word-spacing
    /// folded in for a U+0020 cluster). `x + advance` is the next glyph's pen
    /// origin; for the trailing glyph it closes the run's true right extent, so
    /// hit-testing and the a11y mirror read the real run width off the glyphs
    /// instead of estimating a uniform trailing advance.
    pub advance: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct RectPrimitive {
    pub x: Number,
    pub y: Number,
    pub w: Number,
    pub h: Number,
    pub fill: String,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct LinePrimitive {
    pub x1: Number,
    pub y1: Number,
    pub x2: Number,
    pub y2: Number,
    pub stroke_width: Number,
    pub color: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<Vec<Number>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub role: Option<LineRole>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_style: Option<DisplayBorderStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_owner: Option<BorderOwner>,
    /// Ownership metadata for table/cell border and cut lines: `attrs.cell`
    /// names the owning grid cell (with its border-ownership flags) and
    /// `attrs.table` the enclosing fragment, replacing the consumer-side
    /// geometric fallback association. Every field is optional/defaulted, so
    /// a default `DocAttrs` serializes to zero extra bytes. `doc_attrs_mut` deliberately keeps
    /// returning `None` for lines so the generic sdt/clip/paragraph stamping
    /// passes stay line-inert; table emission stamps these attrs explicitly.
    #[serde(flatten, default)]
    pub attrs: DocAttrs,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderOwner {
    #[serde(rename = "cell")]
    Cell,
    #[serde(rename = "fragment")]
    Fragment,
    #[serde(rename = "paragraph")]
    Paragraph,
    #[serde(rename = "textBox")]
    TextBox,
}

impl LinePrimitive {
    fn contract_defaults() -> Self {
        Self {
            x1: px(0.0),
            y1: px(0.0),
            x2: px(0.0),
            y2: px(0.0),
            stroke_width: px(0.0),
            color: String::new(),
            dash: None,
            role: None,
            border_style: None,
            secondary_color: None,
            opacity: None,
            border_owner: None,
            attrs: DocAttrs::default(),
        }
    }
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum DisplayBorderStyle {
    #[serde(rename = "solid")]
    Solid,
    #[serde(rename = "double")]
    Double,
    #[serde(rename = "dotted")]
    Dotted,
    #[serde(rename = "dashed")]
    Dashed,
    #[serde(rename = "dashDot")]
    DashDot,
    #[serde(rename = "dashDotDot")]
    DashDotDot,
    #[serde(rename = "triple")]
    Triple,
    #[serde(rename = "thinThick")]
    ThinThick,
    #[serde(rename = "thickThin")]
    ThickThin,
    #[serde(rename = "wave")]
    Wave,
    #[serde(rename = "doubleWave")]
    DoubleWave,
    #[serde(rename = "groove")]
    Groove,
    #[serde(rename = "ridge")]
    Ridge,
    #[serde(rename = "inset")]
    Inset,
    #[serde(rename = "outset")]
    Outset,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq)]
pub enum LineRole {
    #[serde(rename = "border")]
    Border,
    #[serde(rename = "table-border")]
    TableBorder,
    #[serde(rename = "table-cut")]
    TableCut,
    #[serde(rename = "separator")]
    Separator,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PageBorderPrimitive {
    pub x: Number,
    pub y: Number,
    pub w: Number,
    pub h: Number,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z_order: Option<PageBorderZOrder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<PageBorderSide>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<PageBorderSide>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<PageBorderSide>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<PageBorderSide>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageBorderZOrder {
    #[serde(rename = "front")]
    Front,
    #[serde(rename = "back")]
    Back,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct PageBorderSide {
    pub width: Number,
    pub color: String,
    pub style: String,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ImagePrimitive {
    pub rel_id: String,
    pub x: Number,
    pub y: Number,
    pub w: Number,
    pub h: Number,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation_deg: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<Number>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filter: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub decorative: bool,
    /// fractions 0..1
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop: Option<Crop>,
    /// alternative text (`wp:docPr` `descr`), capped at
    /// [`MAX_ALT_TEXT_CHARS`] — file-derived, attacker-controlled data
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct ContentFrame {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<Number>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<Number>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ShapePrimitive {
    pub x: Number,
    pub y: Number,
    pub w: Number,
    pub h: Number,
    pub geometry_path: Vec<ShapePathCommand>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke: Option<ShapeStrokePrimitive>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<ShapeTransformPrimitive>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub decorative: bool,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(tag = "type")]
pub enum ShapePathCommand {
    #[serde(rename = "move")]
    Move { x: Number, y: Number },
    #[serde(rename = "line")]
    Line { x: Number, y: Number },
    #[serde(rename = "quad")]
    Quad {
        cpx: Number,
        cpy: Number,
        x: Number,
        y: Number,
    },
    #[serde(rename = "cubic")]
    Cubic {
        cp1x: Number,
        cp1y: Number,
        cp2x: Number,
        cp2y: Number,
        x: Number,
        y: Number,
    },
    #[serde(rename = "close")]
    Close,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStrokePrimitive {
    pub color: String,
    pub width: Number,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTransformPrimitive {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation: Option<Number>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub flip_h: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub flip_v: bool,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct Crop {
    pub top: Number,
    pub right: Number,
    pub bottom: Number,
    pub left: Number,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct DecorationPrimitive {
    pub deco: DecoKind,
    pub x: Number,
    pub y: Number,
    pub w: Number,
    pub h: Number,
    pub color: String,
    /// Dashed rule instead of a solid one, used for the tracked-change
    /// insertion underline. Omitted for a plain solid decoration.
    #[serde(default, skip_serializing_if = "is_false")]
    pub dashed: bool,
    /// dotted rule instead of a solid one — set for hidden-run dotted underline.
    #[serde(default, skip_serializing_if = "is_false")]
    pub dotted: bool,
    #[serde(flatten)]
    pub attrs: DocAttrs,
}

/// serde `skip_serializing_if` predicate: omit `false` bools from the wire form.
fn is_false(b: &bool) -> bool {
    !*b
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq)]
pub enum DecoKind {
    #[serde(rename = "underline")]
    Underline,
    #[serde(rename = "strike")]
    Strike,
    #[serde(rename = "highlight")]
    Highlight,
    #[serde(rename = "comment-range")]
    CommentRange,
    #[serde(rename = "spell")]
    Spell,
}

/// The parsed build envelope: the measured blocks and options pagination also
/// saw, the `Layout` it produced, and the display-only extras.
pub struct BuildInput {
    contract_version: Option<u32>,
    measured: Vec<MeasuredBlockIn>,
    options: Value,
    layout: LayoutIn,
    /// Optional header/footer payload; see [`crate::hf_bands`] for its shape.
    /// Absent means the pages carry no bands.
    headers_footers: Option<crate::hf_bands::HeadersFootersIn>,
    headers_footers_content: Option<HeadersFootersContentIn>,
    /// Font fallback chains keyed `"<family lowercase>|<bold>|<italic>"`, the
    /// same map measurement used. Resolving these against a populated font
    /// store is what gates glyph-run emission.
    font_chains: HashMap<String, Vec<u32>>,
    resolved_comment_ids: Vec<i64>,
    comment_authors: Vec<CommentAuthorIn>,
    comment_threads: Vec<CommentThreadIn>,
}

/// Parsed display input retained by the editing engine across edits. Its fields
/// stay private so the display-input contract can evolve without becoming a
/// second public layout model.
pub struct ResidentDisplayInput {
    input: BuildInput,
}

impl std::fmt::Debug for ResidentDisplayInput {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ResidentDisplayInput")
            .field("measured_blocks", &self.input.measured.len())
            .field("pages", &self.input.layout.pages.len())
            .finish_non_exhaustive()
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct BuildInputWire {
    #[serde(default)]
    contract_version: Option<u32>,
    #[serde(default)]
    measured: Vec<MeasuredBlockIn>,
    #[serde(default)]
    options: Value,
    layout: LayoutIn,
    #[serde(default)]
    headers_footers: Option<Value>,
    #[serde(default)]
    font_chains: HashMap<String, Vec<u32>>,
    #[serde(default)]
    resolved_comment_ids: Vec<i64>,
    #[serde(default)]
    comment_authors: Vec<CommentAuthorIn>,
    #[serde(default)]
    comment_threads: Vec<CommentThreadIn>,
}

impl<'de> Deserialize<'de> for BuildInput {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let wire = BuildInputWire::deserialize(deserializer)?;
        let headers_footers = wire
            .headers_footers
            .clone()
            .map(serde_json::from_value)
            .transpose()
            .map_err(serde::de::Error::custom)?;
        let headers_footers_content = wire
            .headers_footers
            .map(serde_json::from_value)
            .transpose()
            .map_err(serde::de::Error::custom)?;
        Ok(Self {
            contract_version: wire.contract_version,
            measured: wire.measured,
            options: wire.options,
            layout: wire.layout,
            headers_footers,
            headers_footers_content,
            font_chains: wire.font_chains,
            resolved_comment_ids: wire.resolved_comment_ids,
            comment_authors: wire.comment_authors,
            comment_threads: wire.comment_threads,
        })
    }
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct HeadersFootersContentIn {
    #[serde(default)]
    header_distance: Option<f64>,
    #[serde(default)]
    footer_distance: Option<f64>,
    #[serde(default)]
    variants: Vec<HfVariantContentIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct HfVariantContentIn {
    r_id: String,
    kind: HfKind,
    #[serde(default)]
    measured: Vec<MeasuredBlockIn>,
    #[serde(default)]
    height: Option<f64>,
    #[serde(default)]
    flow_height: Option<f64>,
    #[serde(default)]
    visual_top: Option<f64>,
    #[serde(default)]
    visual_bottom: Option<f64>,
    #[serde(default)]
    field_widths: Vec<HfFieldWidthContentIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct HfFieldWidthContentIn {
    pm_start: i64,
    fallback_width: f64,
    #[serde(default)]
    per_page: Vec<f64>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct CommentAuthorIn {
    #[serde(default)]
    id: Option<String>,
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    palette_index: Option<u64>,
    #[serde(default)]
    color: Option<String>,
}

/// One comment thread from the display-list envelope.
#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct CommentThreadIn {
    #[serde(default)]
    id: Option<i64>,
    #[serde(default)]
    author_id: Option<String>,
    #[serde(default)]
    author_name: Option<String>,
    #[serde(default)]
    date: Option<String>,
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    replies: Vec<CommentReplyIn>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct CommentReplyIn {
    #[serde(default)]
    author_name: Option<String>,
    #[serde(default)]
    date: Option<String>,
    #[serde(default)]
    text: Option<String>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct RenderOptionsIn {
    #[serde(default)]
    page_borders: Option<PageBordersIn>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct PageBordersIn {
    #[serde(default)]
    top: Option<PageBorderSpecIn>,
    #[serde(default)]
    right: Option<PageBorderSpecIn>,
    #[serde(default)]
    bottom: Option<PageBorderSpecIn>,
    #[serde(default)]
    left: Option<PageBorderSpecIn>,
    #[serde(default)]
    display: Option<String>,
    #[serde(default)]
    offset_from: Option<String>,
    #[serde(default)]
    z_order: Option<String>,
}

#[derive(Deserialize, Clone)]
struct PageBorderSpecIn {
    #[serde(default)]
    style: Option<String>,
    #[serde(default)]
    color: Option<Value>,
    /// width in eighths of a point (OOXML w:sz)
    #[serde(default)]
    size: Option<f64>,
    /// spacing from page/text in points
    #[serde(default)]
    space: Option<f64>,
}

#[derive(Deserialize, Clone)]
#[serde(tag = "kind")]
pub(crate) enum WatermarkIn {
    #[serde(rename = "text")]
    Text(TextWatermarkIn),
    #[serde(rename = "picture")]
    Picture(PictureWatermarkIn),
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextWatermarkIn {
    #[serde(default)]
    text: String,
    #[serde(default)]
    font: Option<String>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    semitransparent: Option<bool>,
    #[serde(default)]
    layout: Option<String>,
    /// points; absent means Word's "Auto" sizing
    #[serde(default)]
    font_size: Option<f64>,
    #[serde(default)]
    decorative: Option<bool>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PictureWatermarkIn {
    /// Parsed package watermarks usually carry a renderable data URL. This is
    /// the same key shape the canvas image resolver accepts for normal images.
    #[serde(default)]
    data_url: Option<String>,
    /// Fallback/testing key when no data URL is available.
    #[serde(default)]
    rel_id: Option<String>,
    #[serde(default)]
    scale: Option<f64>,
    #[serde(default)]
    washout: Option<bool>,
    #[serde(default)]
    width_emu: Option<f64>,
    #[serde(default)]
    height_emu: Option<f64>,
    #[serde(default)]
    decorative: Option<bool>,
}

#[derive(Deserialize)]
pub(crate) struct MeasuredBlockIn {
    pub(crate) block: BlockIn,
    pub(crate) measure: MeasureIn,
}

// variants boxed: a transient deserialization mirror, but the enum nests
// inside cell block lists so the size difference would multiply
#[derive(Deserialize)]
#[serde(tag = "kind")]
pub(crate) enum BlockIn {
    #[serde(rename = "paragraph")]
    Paragraph(Box<ParagraphBlockIn>),
    #[serde(rename = "table")]
    Table(Box<TableBlockIn>),
    #[serde(rename = "image")]
    Image(Box<ImageBlockIn>),
    #[serde(rename = "textBox")]
    TextBox(Box<TextBoxBlockIn>),
    #[serde(rename = "shape")]
    Shape(Box<ShapeBlockIn>),
    #[serde(rename = "chart")]
    Chart(Box<ChartBlockIn>),
    #[serde(other)]
    Unsupported,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ParagraphBlockIn {
    #[serde(default)]
    sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    /// stable Word `w14:paraId`, threaded onto the paragraph's
    /// primitives so the a11y mirror can emit `data-para-id`
    #[serde(default)]
    pub(crate) para_id: Option<String>,
    #[serde(default)]
    runs: Vec<RunIn>,
    #[serde(default)]
    pub(crate) attrs: Option<ParaAttrsIn>,
    #[serde(default)]
    pub(crate) pm_start: Option<i64>,
    #[serde(default)]
    pub(crate) pm_end: Option<i64>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SdtGroupIn {
    #[serde(default)]
    id: String,
    #[serde(default)]
    sdt_type: String,
    #[serde(default)]
    tag: Option<String>,
    #[serde(default)]
    alias: Option<String>,
    #[serde(default)]
    lock: Option<String>,
    #[serde(default)]
    checked: Option<bool>,
    #[serde(default)]
    bound: Option<bool>,
    #[serde(default)]
    repeating_item: Option<bool>,
}

#[derive(Deserialize, Clone)]
#[serde(tag = "kind")]
enum RunIn {
    #[serde(rename = "text")]
    Text(TextRunIn),
    #[serde(rename = "tab")]
    Tab(TabRunIn),
    #[serde(rename = "image")]
    Image(ImageRunIn),
    #[serde(rename = "lineBreak")]
    LineBreak(LineBreakRunIn),
    #[serde(rename = "field")]
    Field(FieldRunIn),
    #[serde(other)]
    Unsupported,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct RunFormattingIn {
    #[serde(default)]
    bold: Option<bool>,
    #[serde(default)]
    italic: Option<bool>,
    /// bool or { style, color }
    #[serde(default)]
    underline: Option<Value>,
    #[serde(default)]
    strike: Option<bool>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    highlight: Option<String>,
    #[serde(default)]
    font_family: Option<String>,
    /// points
    #[serde(default)]
    font_size: Option<f64>,
    #[serde(default)]
    letter_spacing: Option<f64>,
    #[serde(default)]
    superscript: Option<bool>,
    #[serde(default)]
    subscript: Option<bool>,
    #[serde(default)]
    all_caps: Option<bool>,
    #[serde(default)]
    small_caps: Option<bool>,
    #[serde(default)]
    position_px: Option<f64>,
    #[serde(default)]
    horizontal_scale: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    kerning_min_pt: Option<f64>,
    #[serde(default)]
    imprint: Option<bool>,
    #[serde(default)]
    emboss: Option<bool>,
    #[serde(default)]
    text_shadow: Option<bool>,
    #[serde(default)]
    text_outline: Option<bool>,
    #[serde(default)]
    emphasis_mark: Option<String>,
    #[serde(default)]
    hidden: Option<bool>,
    #[serde(default)]
    rtl: Option<bool>,
    #[serde(default)]
    text_effect: Option<String>,
    /// modern w14 text effects payload, passed through verbatim to the
    /// text/glyph primitives (the canvas backend owns the paint recipe)
    #[serde(default)]
    modern_effects: Option<Value>,
    #[serde(default)]
    footnote_ref_id: Option<i64>,
    #[serde(default)]
    endnote_ref_id: Option<i64>,
    #[serde(default)]
    comment_ids: Option<Vec<i64>>,
    #[serde(default)]
    is_insertion: Option<bool>,
    #[serde(default)]
    is_deletion: Option<bool>,
    #[serde(default)]
    change_author: Option<String>,
    #[serde(default)]
    change_date: Option<String>,
    #[serde(default)]
    change_revision_id: Option<i64>,
    #[serde(default)]
    hyperlink: Option<HyperlinkIn>,
    #[serde(default)]
    inline_sdt_widget: Option<InlineSdtWidgetAttrs>,
    #[serde(default)]
    language: Option<RunLanguageIn>,
    #[serde(default)]
    logical_order: Option<u64>,
    #[serde(default)]
    bidi_level: Option<u8>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct RunLanguageIn {
    #[serde(default)]
    latin: Option<String>,
    #[serde(default)]
    east_asia: Option<String>,
    #[serde(default)]
    bidi: Option<String>,
}

#[derive(Deserialize, Clone, Default, Debug, PartialEq)]
#[serde(rename_all = "camelCase")]
struct RevisionInfoIn {
    #[serde(default)]
    revision_id: Option<i64>,
    #[serde(default)]
    author: Option<String>,
    #[serde(default)]
    date: Option<String>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct HyperlinkIn {
    #[serde(default)]
    href: Option<String>,
    #[serde(default)]
    no_default_style: Option<bool>,
    #[serde(default)]
    tooltip: Option<String>,
    #[serde(default)]
    target: Option<String>,
    #[serde(default)]
    history: Option<bool>,
    #[serde(default)]
    doc_location: Option<String>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct TextRunIn {
    #[serde(default)]
    text: String,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    #[serde(flatten)]
    fmt: RunFormattingIn,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct TabRunIn {
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    /// resolved advance in px, filled in by measurement
    #[serde(default)]
    width: Option<f64>,
    #[serde(default)]
    leader_glyphs: Option<LeaderGlyphIn>,
    #[serde(flatten)]
    fmt: RunFormattingIn,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct LeaderGlyphIn {
    #[serde(default)]
    glyph: Option<String>,
    #[serde(default)]
    count: Option<u64>,
    #[serde(default)]
    advance: Option<f64>,
    #[serde(default)]
    font: Option<String>,
    #[serde(default)]
    font_size: Option<f64>,
    #[serde(default)]
    color: Option<String>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ImageRunIn {
    #[serde(default)]
    src: String,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    /// Alternative text from `wp:docPr` `descr`.
    #[serde(default)]
    alt: Option<String>,
    #[serde(default)]
    transform: Option<String>,
    #[serde(default)]
    wrap_type: Option<String>,
    #[serde(default)]
    display_mode: Option<String>,
    /// CSS float direction (`left`/`right`/`none`) — resolves the horizontal
    /// anchor for a float that carries no explicit `position` (imageWrapText /
    /// resolveHorizontalAnchor).
    #[serde(default)]
    css_float: Option<String>,
    /// anchor position for a floating image run (`wp:positionH`/`wp:positionV`),
    /// resolved to a page rect by [`resolve_anchored_position`]
    #[serde(default)]
    position: Option<AnchorPosIn>,
    #[serde(default)]
    crop_top: Option<f64>,
    #[serde(default)]
    crop_right: Option<f64>,
    #[serde(default)]
    crop_bottom: Option<f64>,
    #[serde(default)]
    crop_left: Option<f64>,
    #[serde(default)]
    opacity: Option<f64>,
    #[serde(default)]
    rotation_deg: Option<f64>,
    #[serde(default)]
    flip_h: Option<bool>,
    #[serde(default)]
    flip_v: Option<bool>,
    #[serde(default)]
    rotation_bounds: Option<RotationBoundsIn>,
    #[serde(default)]
    effects: Vec<Value>,
    #[serde(default)]
    outline: Option<Value>,
    #[serde(default)]
    decorative: Option<bool>,
    #[serde(default)]
    hyperlink: Option<HyperlinkIn>,
    #[serde(default)]
    is_insertion: Option<bool>,
    #[serde(default)]
    is_deletion: Option<bool>,
    #[serde(default)]
    change_author: Option<String>,
    #[serde(default)]
    change_date: Option<String>,
    #[serde(default)]
    change_revision_id: Option<f64>,
    #[serde(default)]
    hlink_href: Option<String>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RotationBoundsIn {
    #[serde(default)]
    width: Option<f64>,
    #[serde(default)]
    height: Option<f64>,
    #[serde(default)]
    offset_x: Option<f64>,
    #[serde(default)]
    offset_y: Option<f64>,
}

/// anchor of a floating image/text-box run (`ImageRunPosition`): one axis each,
/// resolved against the page geometry in [`resolve_anchored_position`].
#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct AnchorPosIn {
    #[serde(default)]
    horizontal: Option<AnchorAxisIn>,
    #[serde(default)]
    vertical: Option<AnchorAxisIn>,
}

/// one axis of an anchor: an OOXML `relativeFrom` band plus either an `align`
/// keyword or a `posOffset` (EMU). Mirrors `ImageRunPosition.{horizontal,vertical}`.
#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct AnchorAxisIn {
    #[serde(default)]
    relative_to: Option<String>,
    /// offset from the band base, in EMU (converted with [`emu_to_px`])
    #[serde(default)]
    pos_offset: Option<f64>,
    #[serde(default)]
    align: Option<String>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct LineBreakRunIn {
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    #[allow(dead_code)]
    pm_end: Option<i64>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FieldRunIn {
    #[serde(default)]
    field_type: Option<String>,
    /// raw Word field type token (a11y identity; never evaluated)
    #[serde(default)]
    raw_type: Option<String>,
    /// raw instruction text (a11y identity; INERT — never parsed/executed)
    #[serde(default)]
    instruction: Option<String>,
    #[serde(default)]
    fallback: Option<String>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    #[serde(flatten)]
    fmt: RunFormattingIn,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ParaAttrsIn {
    #[serde(default)]
    alignment: Option<String>,
    /// Resolved paragraph spacing.
    #[serde(default)]
    pub(crate) spacing: Option<SpacingIn>,
    #[serde(default)]
    bidi: Option<bool>,
    #[serde(default)]
    shading: Option<String>,
    #[serde(default)]
    indent: Option<IndentIn>,
    #[serde(default)]
    borders: Option<ParaBordersIn>,
    /// Pre-computed list marker text.
    #[serde(default)]
    list_marker: Option<String>,
    /// w:vanish on the numbering level rPr — a hidden marker reserves no slot
    #[serde(default)]
    list_marker_hidden: Option<bool>,
    #[serde(default)]
    list_marker_font_family: Option<String>,
    #[serde(default)]
    list_marker_font_size: Option<f64>,
    #[serde(default)]
    list_marker_revision: Option<RevisionKind>,
    #[serde(default)]
    default_font_family: Option<String>,
    #[serde(default)]
    default_font_size: Option<f64>,
    #[serde(default)]
    tabs: Option<Vec<TabStopIn>>,
    /// `<w:pPr><w:rPr><w:ins/>`: tracked insertion on the paragraph mark.
    #[serde(default)]
    p_pr_ins: Option<RevisionInfoIn>,
    /// `<w:pPr><w:rPr><w:del/>`: tracked deletion on the paragraph mark.
    #[serde(default)]
    p_pr_del: Option<RevisionInfoIn>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct TabStopIn {
    #[serde(default)]
    val: Option<String>,
    #[serde(default)]
    pos: Option<f64>,
    #[serde(default)]
    leader: Option<String>,
}

#[derive(Deserialize, Default, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SpacingIn {
    #[serde(default)]
    pub(crate) before: Option<f64>,
    #[serde(default)]
    pub(crate) after: Option<f64>,
}

#[derive(Deserialize, Default, Clone, Copy)]
#[serde(rename_all = "camelCase")]
struct IndentIn {
    #[serde(default)]
    left: Option<f64>,
    #[serde(default)]
    right: Option<f64>,
    #[serde(default)]
    first_line: Option<f64>,
    #[serde(default)]
    hanging: Option<f64>,
}

#[derive(Deserialize, Default, Clone, PartialEq)]
pub(crate) struct ParaBordersIn {
    #[serde(default)]
    top: Option<BorderEdgeIn>,
    #[serde(default)]
    bottom: Option<BorderEdgeIn>,
    #[serde(default)]
    left: Option<BorderEdgeIn>,
    #[serde(default)]
    right: Option<BorderEdgeIn>,
    #[serde(default)]
    between: Option<BorderEdgeIn>,
    #[serde(default)]
    bar: Option<BorderEdgeIn>,
}

#[derive(Deserialize, Default, Clone, PartialEq)]
struct BorderEdgeIn {
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    style: Option<String>,
    #[serde(default)]
    width: Option<f64>,
    #[serde(default)]
    space: Option<f64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableBlockIn {
    #[serde(default)]
    sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    #[serde(default)]
    pub(crate) rows: Vec<TableRowIn>,
    #[serde(default)]
    bidi: Option<bool>,
    #[serde(default)]
    justification: Option<String>,
    #[serde(default)]
    indent: Option<f64>,
    #[serde(default)]
    caption: Option<String>,
    #[serde(default)]
    description: Option<String>,
    /// `<w:tblpPr>` placement. A floating table does not advance the
    /// header/footer flow cursor.
    #[serde(default)]
    pub(crate) floating: Option<FloatingTablePositionIn>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FloatingTablePositionIn {
    #[serde(default)]
    pub(crate) horz_anchor: Option<String>,
    #[serde(default)]
    pub(crate) tblp_x: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) tblp_x_spec: Option<String>,
    #[serde(default)]
    pub(crate) vert_anchor: Option<String>,
    #[serde(default)]
    pub(crate) tblp_y: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) tblp_y_spec: Option<String>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) top_from_text: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) right_from_text: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) bottom_from_text: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) left_from_text: Option<f64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableRowIn {
    #[serde(default)]
    cells: Vec<TableCellIn>,
    #[serde(default)]
    is_header: Option<bool>,
    #[serde(default)]
    tracked_ins: Option<RevisionInfoIn>,
    #[serde(default)]
    tracked_del: Option<RevisionInfoIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct TableCellIn {
    #[serde(default)]
    blocks: Vec<BlockIn>,
    #[serde(default)]
    col_span: Option<u32>,
    #[serde(default)]
    row_span: Option<u32>,
    #[serde(default)]
    background: Option<String>,
    #[serde(default)]
    borders: Option<CellBordersIn>,
    #[serde(default)]
    padding: Option<CellPaddingIn>,
    /// Cell content vertical alignment.
    #[serde(default)]
    vertical_align: Option<String>,
    #[serde(default)]
    no_wrap: Option<bool>,
    #[serde(default)]
    tracked_marker: Option<CellMarkerIn>,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct CellMarkerIn {
    kind: StructuralRevisionKind,
    info: RevisionInfoIn,
}

#[derive(Deserialize, Default)]
struct CellBordersIn {
    #[serde(default)]
    top: Option<BorderEdgeIn>,
    #[serde(default)]
    right: Option<BorderEdgeIn>,
    #[serde(default)]
    bottom: Option<BorderEdgeIn>,
    #[serde(default)]
    left: Option<BorderEdgeIn>,
}

#[derive(Deserialize, Default, Clone, Copy)]
struct CellPaddingIn {
    #[serde(default)]
    top: Option<f64>,
    #[serde(default)]
    right: Option<f64>,
    #[serde(default)]
    bottom: Option<f64>,
    #[serde(default)]
    left: Option<f64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ImageBlockIn {
    #[serde(default)]
    pub(crate) sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    #[serde(default)]
    pub(crate) src: String,
    #[serde(default)]
    pub(crate) width: f64,
    #[serde(default)]
    pub(crate) height: f64,
    /// Alternative text from `wp:docPr` `descr`.
    #[serde(default)]
    pub(crate) alt: Option<String>,
    #[serde(default)]
    pub(crate) transform: Option<String>,
    #[serde(default)]
    pub(crate) opacity: Option<f64>,
    #[serde(default)]
    pub(crate) rotation_deg: Option<f64>,
    #[serde(default)]
    pub(crate) flip_h: Option<bool>,
    #[serde(default)]
    pub(crate) flip_v: Option<bool>,
    #[serde(default)]
    pub(crate) rotation_bounds: Option<RotationBoundsIn>,
    #[serde(default)]
    pub(crate) hlink_href: Option<String>,
    #[serde(default)]
    pub(crate) hlink_title: Option<String>,
    #[serde(default)]
    pub(crate) decorative: Option<bool>,
    #[serde(default)]
    pub(crate) crop: Option<CropIn>,
    #[serde(default)]
    pub(crate) effects: Vec<Value>,
    #[serde(default)]
    pub(crate) outline: Option<Value>,
    #[serde(default)]
    pub(crate) pm_start: Option<i64>,
    #[serde(default)]
    pub(crate) pm_end: Option<i64>,
}

#[derive(Deserialize, Clone, Default)]
pub(crate) struct CropIn {
    #[serde(default)]
    top: Option<f64>,
    #[serde(default)]
    right: Option<f64>,
    #[serde(default)]
    bottom: Option<f64>,
    #[serde(default)]
    left: Option<f64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ShapeBlockIn {
    #[serde(default)]
    sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    #[serde(default)]
    #[allow(dead_code)]
    shape_type: Option<String>,
    #[serde(default)]
    geometry_path: Vec<ShapePathCommand>,
    #[serde(default)]
    fill: Option<ShapeFillIn>,
    #[serde(default)]
    stroke: Option<ShapeStrokeIn>,
    #[serde(default)]
    transform: Option<ShapeTransformIn>,
    #[serde(default)]
    #[allow(dead_code)]
    width: f64,
    #[serde(default)]
    #[allow(dead_code)]
    height: f64,
    #[serde(default)]
    #[allow(dead_code)]
    x: Option<f64>,
    #[serde(default)]
    #[allow(dead_code)]
    y: Option<f64>,
    #[serde(default)]
    inner_text: Vec<ParagraphBlockIn>,
    #[serde(default)]
    inner_measures: Vec<ParagraphExtentIn>,
    #[serde(default)]
    children: Vec<ShapeBlockIn>,
    #[serde(default)]
    scene: Option<Value>,
    #[serde(default)]
    effects: Vec<Value>,
    #[serde(default)]
    effect_extent: Option<Value>,
    #[serde(default)]
    text_body_properties: Option<Value>,
    #[serde(default)]
    relative_height: Option<u64>,
    #[serde(default)]
    decorative: Option<bool>,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    description: Option<String>,
    #[serde(default)]
    hidden: Option<bool>,
    #[serde(default)]
    doc_start: Option<i64>,
    #[serde(default)]
    doc_end: Option<i64>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ChartBlockIn {
    #[serde(default)]
    sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    #[serde(default)]
    chart: ChartIn,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    #[serde(default)]
    doc_start: Option<i64>,
    #[serde(default)]
    doc_end: Option<i64>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartIn {
    #[serde(default)]
    chart_type: String,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    series: Vec<ChartSeriesIn>,
    #[serde(default)]
    legend: Option<ChartLegendIn>,
    #[serde(default)]
    axes: Option<ChartAxesIn>,
    #[serde(default)]
    description: Option<String>,
    #[serde(default)]
    decorative: Option<bool>,
    #[serde(default)]
    plot_groups: Vec<ChartPlotGroupIn>,
    #[serde(default)]
    axis_list: Vec<ChartAxisIn>,
    #[serde(default)]
    text: Option<ChartTextIn>,
    #[serde(default)]
    title_text: Option<ChartTextIn>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartTextIn {
    #[serde(default)]
    font: Option<String>,
    #[serde(default)]
    size_pt: Option<f64>,
    #[serde(default)]
    bold: Option<bool>,
    #[serde(default)]
    italic: Option<bool>,
    #[serde(default)]
    color: Option<String>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartPlotGroupIn {
    #[serde(default)]
    chart_type: Option<String>,
    #[serde(default)]
    grouping: Option<String>,
    #[serde(default)]
    series: Vec<ChartSeriesIn>,
    #[serde(default)]
    overlap: Option<f64>,
    #[serde(default)]
    gap_width: Option<f64>,
    #[serde(default)]
    hole_size: Option<f64>,
    #[serde(default)]
    first_slice_angle: Option<f64>,
    #[serde(default)]
    vary_colors: bool,
    #[serde(default)]
    scatter_style: Option<String>,
    #[serde(default)]
    radar_style: Option<String>,
    #[serde(default)]
    bubble_scale: Option<f64>,
    #[serde(default)]
    size_represents: Option<String>,
    #[serde(default)]
    wireframe: Option<bool>,
    #[serde(default)]
    hi_low_lines: bool,
    #[serde(default)]
    up_down_bars: bool,
    #[serde(default)]
    marker: Option<bool>,
    #[serde(default)]
    axis_ids: Vec<String>,
    #[serde(default)]
    data_labels: Option<ChartDataLabelsIn>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartDataLabelsIn {
    #[serde(default)]
    delete: Option<bool>,
    #[serde(default)]
    show_value: Option<bool>,
    #[serde(default)]
    show_category_name: Option<bool>,
    #[serde(default)]
    show_series_name: Option<bool>,
    #[serde(default)]
    show_percent: Option<bool>,
    #[serde(default)]
    show_legend_key: Option<bool>,
    #[serde(default)]
    show_bubble_size: Option<bool>,
    #[serde(default)]
    separator: Option<String>,
    #[serde(default)]
    position: Option<String>,
    #[serde(default)]
    number_format: Option<String>,
    #[serde(default)]
    text: Option<ChartTextIn>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartSeriesIn {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    categories: Vec<String>,
    #[serde(default)]
    values: Vec<f64>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    points: Vec<ChartPointIn>,
    #[serde(default)]
    grouping: Option<String>,
    #[serde(default)]
    marker: Option<Value>,
    #[serde(default)]
    x_values: Vec<f64>,
    #[serde(default)]
    bubble_sizes: Vec<f64>,
    #[serde(default)]
    smooth: bool,
    #[serde(default)]
    data_labels: Option<ChartDataLabelsIn>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartPointIn {
    #[serde(default)]
    index: Option<usize>,
    #[serde(default)]
    value: Option<f64>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    marker: Option<Value>,
    #[serde(default)]
    label: Option<String>,
    #[serde(default)]
    explosion: Option<f64>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartLegendIn {
    #[serde(default)]
    position: Option<String>,
    #[serde(default)]
    visible: Option<bool>,
    #[serde(default)]
    text: Option<ChartTextIn>,
}

#[derive(Deserialize, Default, Clone)]
struct ChartAxesIn {
    #[serde(default)]
    value: Option<ChartAxisIn>,
}

#[derive(Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
struct ChartAxisIn {
    #[serde(default)]
    id: Option<String>,
    #[serde(default)]
    min: Option<f64>,
    #[serde(default)]
    max: Option<f64>,
    #[serde(default)]
    axis_type: Option<String>,
    #[serde(default)]
    logarithmic_base: Option<f64>,
    #[serde(default)]
    reversed: bool,
    #[serde(default)]
    major_unit: Option<f64>,
    #[serde(default)]
    minor_unit: Option<f64>,
    #[serde(default)]
    major_tick_mark: Option<String>,
    #[serde(default)]
    minor_tick_mark: Option<String>,
    #[serde(default)]
    major_gridlines: bool,
    #[serde(default)]
    minor_gridlines: bool,
    #[serde(default)]
    number_format: Option<String>,
    #[serde(default)]
    position: Option<String>,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    hidden: bool,
    #[serde(default)]
    text: Option<ChartTextIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeFillIn {
    #[serde(rename = "type", default)]
    kind: Option<String>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    gradient_type: Option<String>,
    #[serde(default)]
    gradient_angle: Option<f64>,
    #[serde(default)]
    gradient_stops: Vec<Value>,
    #[serde(default)]
    pattern_preset: Option<String>,
    #[serde(default)]
    foreground_color: Option<String>,
    #[serde(default)]
    background_color: Option<String>,
    #[serde(default)]
    picture_rel_id: Option<String>,
    /// resolved SAFE embedded picture source (`data:`/`blob:` minted by the
    /// parser from embedded parts; never an external target)
    #[serde(default)]
    picture_src: Option<String>,
    #[serde(default)]
    picture_src_rect: Option<Value>,
    #[serde(default)]
    picture_fill_mode: Option<String>,
    #[serde(default)]
    picture_tile: Option<Value>,
    #[serde(default)]
    picture_stretch_rect: Option<Value>,
    #[serde(default)]
    picture_opacity: Option<f64>,
    #[serde(default)]
    theme_ref_index: Option<u64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeStrokeIn {
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    width: Option<f64>,
    #[serde(default)]
    dash: Option<String>,
    #[serde(default)]
    compound: Option<String>,
    #[serde(default)]
    alignment: Option<String>,
    #[serde(default)]
    cap: Option<String>,
    #[serde(default)]
    join: Option<String>,
    #[serde(default)]
    miter_limit: Option<f64>,
    #[serde(default)]
    custom_dash: Vec<f64>,
    #[serde(default)]
    head_end: Option<Value>,
    #[serde(default)]
    tail_end: Option<Value>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeTransformIn {
    #[serde(default)]
    rotation: Option<f64>,
    #[serde(default)]
    flip_h: Option<bool>,
    #[serde(default)]
    flip_v: Option<bool>,
}

/// Positioned text-box input.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextBoxBlockIn {
    #[serde(default)]
    sdt_groups: Vec<SdtGroupIn>,
    pub(crate) id: Value,
    #[serde(default)]
    fill_color: Option<String>,
    #[serde(default)]
    outline_width: Option<f64>,
    #[serde(default)]
    outline_color: Option<String>,
    #[serde(default)]
    #[allow(dead_code)]
    outline_style: Option<String>,
    /// internal padding; absent ⇒ [`DEFAULT_TEXTBOX_MARGINS`]
    #[serde(default)]
    margins: Option<TextBoxMarginsIn>,
    /// inner paragraph blocks, index-aligned with the measure's `innerMeasures`
    #[serde(default)]
    content: Vec<ParagraphBlockIn>,
    #[serde(default)]
    display_mode: Option<String>,
    #[serde(default)]
    css_float: Option<String>,
    #[serde(default)]
    wrap_type: Option<String>,
    #[serde(default)]
    position: Option<AnchorPosIn>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    // NB: the box's own pmStart/pmEnd are carried by the TextBoxFragment (read
    // there); the block's copies are ignored (serde drops the unread fields).
}

#[derive(Deserialize, Default, Clone, Copy)]
#[serde(rename_all = "camelCase")]
struct TextBoxMarginsIn {
    #[serde(default)]
    top: f64,
    /// Parsed but unused bottom padding.
    #[serde(default)]
    #[allow(dead_code)]
    bottom: f64,
    #[serde(default)]
    left: f64,
    #[serde(default)]
    right: f64,
}

/// OOXML default text-box margins in pixels.
const DEFAULT_TEXTBOX_MARGINS: TextBoxMarginsIn = TextBoxMarginsIn {
    top: 4.0,
    bottom: 4.0,
    left: 7.0,
    right: 7.0,
};

#[derive(Deserialize)]
#[serde(tag = "kind")]
pub(crate) enum MeasureIn {
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphExtentIn),
    #[serde(rename = "table")]
    Table(TableExtentIn),
    #[serde(rename = "image")]
    Image(ImageExtentIn),
    #[serde(rename = "textBox")]
    TextBox(TextBoxExtentIn),
    #[serde(rename = "shape")]
    Shape(BoxExtentIn),
    #[serde(rename = "chart")]
    Chart(BoxExtentIn),
    #[serde(other)]
    Unsupported,
}

/// text-box measure (mirrors `TextBoxExtent`): the box's resolved size plus its
/// inner paragraphs pre-measured, index-aligned with `TextBoxBlockIn.content`.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextBoxExtentIn {
    #[serde(default)]
    #[allow(dead_code)]
    width: f64,
    /// resolved box height; used by HF stacked-height fallback (`hf_bands`)
    #[serde(default)]
    pub(crate) height: f64,
    #[serde(default)]
    inner_measures: Vec<ParagraphExtentIn>,
}

#[derive(Deserialize, Default, Clone, Copy)]
pub(crate) struct BoxExtentIn {
    #[serde(default)]
    pub(crate) width: f64,
    #[serde(default)]
    pub(crate) height: f64,
}

#[derive(Deserialize, Default, Clone, Copy)]
pub(crate) struct ImageExtentIn {
    #[serde(default)]
    pub(crate) width: f64,
    #[serde(default)]
    pub(crate) height: f64,
}

#[derive(Deserialize)]
pub(crate) struct ParagraphExtentIn {
    #[serde(default)]
    pub(crate) lines: Vec<LineIn>,
    #[serde(rename = "totalHeight", default)]
    pub(crate) total_height: f64,
}

#[derive(Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct LineIn {
    #[serde(default)]
    head_run: usize,
    #[serde(default)]
    head_char: usize,
    #[serde(default)]
    tail_run: usize,
    #[serde(default)]
    tail_char: usize,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    ascent: f64,
    #[serde(default)]
    descent: f64,
    #[serde(default)]
    line_height: f64,
    #[serde(default)]
    synthetic_fallback: bool,
    #[serde(default)]
    left_offset: Option<f64>,
    #[serde(default)]
    right_offset: Option<f64>,
    #[serde(default)]
    float_skip_before: Option<f64>,
    #[serde(default)]
    run_advances: Vec<TypesetRunAdvanceIn>,
    #[serde(default)]
    cluster_advances: Vec<TypesetClusterAdvanceIn>,
    #[serde(default)]
    bidi_slices: Vec<TypesetBidiSliceIn>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct TypesetRunAdvanceIn {
    #[serde(default)]
    run_index: Option<usize>,
    #[serde(default)]
    start_char: Option<usize>,
    #[serde(default)]
    end_char: Option<usize>,
    #[serde(default)]
    advance: Option<f64>,
    #[serde(default)]
    logical_order: Option<u64>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct TypesetClusterAdvanceIn {
    #[serde(default)]
    run_index: Option<usize>,
    #[serde(default)]
    start_char: Option<usize>,
    #[serde(default)]
    end_char: Option<usize>,
    #[serde(default)]
    advance: Option<f64>,
    #[serde(default)]
    x_offset: Option<f64>,
    #[serde(default)]
    bidi_level: Option<u8>,
    #[serde(default)]
    logical_order: Option<u64>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct TypesetBidiSliceIn {
    #[serde(default)]
    run_index: Option<usize>,
    #[serde(default)]
    start_char: Option<usize>,
    #[serde(default)]
    end_char: Option<usize>,
    #[serde(default)]
    advance: Option<f64>,
    #[serde(default)]
    bidi_level: Option<u8>,
    #[serde(default)]
    visual_order: Option<u64>,
    #[serde(default)]
    logical_order: Option<u64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableExtentIn {
    #[serde(default)]
    pub(crate) rows: Vec<TableRowExtentIn>,
    #[serde(default)]
    column_widths: Vec<f64>,
    #[serde(default)]
    total_width: f64,
    #[serde(default)]
    pub(crate) total_height: f64,
}

#[derive(Deserialize)]
pub(crate) struct TableRowExtentIn {
    #[serde(default)]
    cells: Vec<TableCellExtentIn>,
    #[serde(default)]
    height: f64,
}

#[derive(Deserialize)]
struct TableCellExtentIn {
    #[serde(default)]
    blocks: Vec<MeasureIn>,
    #[serde(default)]
    #[allow(dead_code)]
    width: f64,
    #[serde(default)]
    height: f64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct LayoutIn {
    #[serde(default)]
    pages: Vec<PageIn>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PageIn {
    #[serde(default)]
    pub(crate) size: SizeIn,
    #[serde(default)]
    pub(crate) margins: MarginsIn,
    /// 1-based page number (canonical layouts carry it; falls back to index+1)
    #[serde(default)]
    pub(crate) number: Option<u64>,
    #[serde(default)]
    pub(crate) page_label: Option<String>,
    #[serde(default)]
    section_id: Option<String>,
    #[serde(default)]
    pub(crate) section_index: Option<u64>,
    #[serde(default)]
    pub(crate) section_page_index: Option<u64>,
    #[serde(default)]
    section_page_number: Option<u64>,
    #[serde(default)]
    pub(crate) header_footer_refs: Option<PageHeaderFooterRefsIn>,
    #[serde(default)]
    background: Option<String>,
    #[serde(default)]
    columns: Option<ColumnLayoutIn>,
    #[serde(default)]
    note_areas: Vec<NoteAreaIn>,
    #[serde(default)]
    fragments: Vec<FragmentIn>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PageHeaderFooterRefsIn {
    #[serde(default)]
    pub(crate) header_default: Option<String>,
    #[serde(default)]
    pub(crate) header_first: Option<String>,
    #[serde(default)]
    pub(crate) header_even: Option<String>,
    #[serde(default)]
    pub(crate) footer_default: Option<String>,
    #[serde(default)]
    pub(crate) footer_first: Option<String>,
    #[serde(default)]
    pub(crate) footer_even: Option<String>,
}

#[derive(Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
struct ColumnLayoutIn {
    #[serde(default)]
    count: usize,
    #[serde(default)]
    gap: f64,
    #[serde(default)]
    equal_width: Option<bool>,
    #[serde(default)]
    separator: Option<bool>,
    #[serde(default)]
    columns: Vec<ColumnSpecIn>,
}

#[derive(Deserialize, Clone, Default)]
struct ColumnSpecIn {
    #[serde(default)]
    width: Option<f64>,
    #[serde(default)]
    space: Option<f64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct NoteAreaIn {
    #[serde(default)]
    kind: Option<String>,
    #[serde(default)]
    y: Option<f64>,
    #[serde(default)]
    height: Option<f64>,
    #[serde(default)]
    columns: Option<u64>,
    #[serde(default)]
    section_id: Option<String>,
    #[serde(default)]
    separator: Option<NoteItemIn>,
    #[serde(default)]
    notes: Vec<NoteItemIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct NoteItemIn {
    #[serde(default)]
    id: Option<i64>,
    #[serde(default)]
    blocks: Vec<BlockIn>,
    #[serde(default)]
    measures: Vec<MeasureIn>,
    #[serde(default)]
    height: Option<f64>,
    #[serde(default)]
    anchor_doc_start: Option<i64>,
    #[serde(default)]
    anchor_doc_end: Option<i64>,
    /// formatted reference label (display number / custom mark); file-derived
    #[serde(default)]
    display_label: Option<String>,
}

#[derive(Deserialize, Default, Clone, Copy)]
pub(crate) struct SizeIn {
    #[serde(default)]
    pub(crate) w: f64,
    #[serde(default)]
    pub(crate) h: f64,
}

/// Page margins as serialized in the `Layout`. `header` and `footer` are the
/// band distances header/footer composition falls back to.
#[derive(Deserialize, Default, Clone, Copy)]
pub(crate) struct MarginsIn {
    #[serde(default)]
    pub(crate) top: f64,
    #[serde(default)]
    pub(crate) right: f64,
    #[serde(default)]
    pub(crate) bottom: f64,
    #[serde(default)]
    pub(crate) left: f64,
    #[serde(default)]
    pub(crate) header: Option<f64>,
    #[serde(default)]
    pub(crate) footer: Option<f64>,
}

#[derive(Deserialize)]
#[serde(tag = "kind")]
enum FragmentIn {
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphFragmentIn),
    #[serde(rename = "table")]
    Table(TableFragmentIn),
    #[serde(rename = "image")]
    Image(ImageFragmentIn),
    #[serde(rename = "textBox")]
    TextBox(TextBoxFragmentIn),
    #[serde(rename = "shape")]
    Shape(ShapeFragmentIn),
    #[serde(rename = "chart")]
    Chart(ChartFragmentIn),
    #[serde(other)]
    Unsupported,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ParagraphFragmentIn {
    pub(crate) block_id: Value,
    #[serde(default)]
    pub(crate) x: f64,
    #[serde(default)]
    pub(crate) y: f64,
    #[serde(default)]
    pub(crate) width: f64,
    #[serde(default)]
    pub(crate) height: f64,
    #[serde(default)]
    pub(crate) from_line: usize,
    #[serde(default)]
    pub(crate) to_line: usize,
    #[serde(default)]
    pub(crate) pm_start: Option<i64>,
    #[serde(default)]
    pub(crate) pm_end: Option<i64>,
    #[serde(default)]
    pub(crate) carried_from_prev: Option<bool>,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) carried_to_next: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableFragmentIn {
    pub(crate) block_id: Value,
    #[serde(default)]
    pub(crate) x: f64,
    #[serde(default)]
    pub(crate) y: f64,
    #[serde(default)]
    #[allow(dead_code)]
    pub(crate) width: f64,
    #[serde(default)]
    pub(crate) height: f64,
    #[serde(default)]
    pub(crate) row_start: usize,
    #[serde(default)]
    pub(crate) row_end: usize,
    #[serde(default)]
    pub(crate) clip_top: Option<f64>,
    #[serde(default)]
    pub(crate) clip_bottom: Option<f64>,
    #[serde(default)]
    pub(crate) header_row_count: Option<usize>,
    #[serde(default)]
    pub(crate) carried_from_prev: Option<bool>,
    #[serde(default)]
    pub(crate) carried_to_next: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ImageFragmentIn {
    block_id: Value,
    #[serde(default)]
    x: f64,
    #[serde(default)]
    y: f64,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct TextBoxFragmentIn {
    block_id: Value,
    #[serde(default)]
    x: f64,
    #[serde(default)]
    y: f64,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    /// Unused stacking hint.
    #[serde(default)]
    #[allow(dead_code)]
    is_floating: Option<bool>,
    #[serde(default)]
    #[allow(dead_code)]
    z_index: Option<i64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeFragmentIn {
    block_id: Value,
    #[serde(default)]
    x: f64,
    #[serde(default)]
    y: f64,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    #[serde(default)]
    doc_start: Option<i64>,
    #[serde(default)]
    doc_end: Option<i64>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
    #[serde(default)]
    #[allow(dead_code)]
    is_anchored: Option<bool>,
    #[serde(default)]
    #[allow(dead_code)]
    z_index: Option<i64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFragmentIn {
    block_id: Value,
    #[serde(default)]
    x: f64,
    #[serde(default)]
    y: f64,
    #[serde(default)]
    width: f64,
    #[serde(default)]
    height: f64,
    #[serde(default)]
    doc_start: Option<i64>,
    #[serde(default)]
    doc_end: Option<i64>,
    #[serde(default)]
    pm_start: Option<i64>,
    #[serde(default)]
    pm_end: Option<i64>,
}

// ---------------------------------------------------------------------------
// numeric canonicalization
// ---------------------------------------------------------------------------

/// canonical JSON number: rounded to 3 decimals (golden precision), -0
/// collapsed, integral values emitted as integers so 96 stays `96`, not `96.0`.
pub fn px(v: f64) -> Number {
    let r = (v * 1000.0).round() / 1000.0;
    let r = if r == 0.0 { 0.0 } else { r };
    if r.fract() == 0.0 && r.abs() < 9.0e15 {
        Number::from(r as i64)
    } else {
        Number::from_f64(r).unwrap_or_else(|| Number::from(0))
    }
}

fn num_f64(n: &Number) -> f64 {
    n.as_f64().unwrap_or(0.0)
}

/// round a coordinate to golden precision (3 decimals, -0 collapsed) as a plain
/// `f64` — GlyphRun glyph positions are `f64` by contract (unlike the `Number`
/// coordinates on the other primitives), so this keeps them deterministic
/// without the integer-collapse `px` applies.
fn round3(v: f64) -> f64 {
    let r = (v * 1000.0).round() / 1000.0;
    if r == 0.0 { 0.0 } else { r }
}

/// Primitive block identity for string or numeric identifiers.
#[derive(Clone, Default, Debug, PartialEq)]
pub(crate) struct BlockRef {
    id: Option<Number>,
    key: Option<String>,
}

impl BlockRef {
    pub(crate) fn of(raw: &Value) -> Self {
        match raw {
            Value::Number(n) => BlockRef {
                id: Some(n.clone()),
                key: None,
            },
            Value::String(s) => BlockRef {
                id: None,
                key: Some(s.clone()),
            },
            _ => BlockRef::default(),
        }
    }

    /// DocAttrs carrying only the block identity; callers fill the rest
    pub(crate) fn attrs(&self) -> DocAttrs {
        DocAttrs {
            block_id: self.id.clone(),
            block_key: self.key.clone(),
            ..Default::default()
        }
    }
}

/// Returns the canonical string key for a block identifier.
fn block_key(id: &Value) -> String {
    match id {
        Value::String(s) => s.clone(),
        other => other.to_string(),
    }
}

const DEFAULT_FONT_PT: f64 = 11.0;
const DEFAULT_FONT_FAMILY: &str = "Calibri";
const REVISION_INS_COLOR: &str = "#2e7d32";
const REVISION_DEL_COLOR: &str = "#c62828";
const REVISION_MERGE_COLOR: &str = "#5f6368";
const STRUCTURAL_CHANGE_BAR_OFFSET_X: f64 = -10.0;
const STRUCTURAL_CHANGE_BAR_WIDTH: f64 = 2.0;
const CELL_STRUCTURAL_BAR_HEIGHT: f64 = 3.0;
const PARAGRAPH_MARK_GLYPH_GAP: f64 = 2.0;
const PARAGRAPH_MARK_GLYPH_WIDTH: f64 = 8.0;

fn font_px_of(fmt: &RunFormattingIn) -> f64 {
    fmt.font_size.unwrap_or(DEFAULT_FONT_PT) * 96.0 / 72.0
}

fn script_scale_of(fmt: &RunFormattingIn) -> f64 {
    if fmt.superscript == Some(true) || fmt.subscript == Some(true) {
        0.75
    } else {
        1.0
    }
}

fn effective_font_px_of(fmt: &RunFormattingIn) -> f64 {
    font_px_of(fmt) * script_scale_of(fmt)
}

fn fallback_scalar_count(text: &str, all_caps: Option<bool>) -> usize {
    if all_caps == Some(true) {
        text.chars().flat_map(char::to_uppercase).count()
    } else {
        text.chars().count()
    }
}

fn fallback_text_width(text: &str, fmt: &RunFormattingIn, default_font_pt: f64) -> f64 {
    let font_size = fmt
        .font_size
        .filter(|size| size.is_finite() && *size > 0.0)
        .unwrap_or(default_font_pt);
    let font_px = font_size * 96.0 / 72.0 * script_scale_of(fmt);
    let letter_spacing = fmt
        .letter_spacing
        .filter(|spacing| spacing.is_finite() && *spacing > 0.0 && *spacing <= 1_000.0)
        .unwrap_or(0.0);
    let horizontal_scale = fmt
        .horizontal_scale
        .filter(|scale| scale.is_finite() && *scale > 0.0 && *scale <= 600.0)
        .map(|scale| scale / 100.0)
        .unwrap_or(1.0);
    fallback_scalar_count(text, fmt.all_caps) as f64 * (font_px + letter_spacing) * horizontal_scale
}

/// Paint-only baseline offset. A positive `positionPx` raises the text, which
/// is a smaller y coordinate.
fn baseline_y_of(fmt: &RunFormattingIn, baseline: f64) -> f64 {
    let mut y = baseline - fmt.position_px.unwrap_or(0.0);
    let script_font = effective_font_px_of(fmt);
    if fmt.superscript == Some(true) {
        y -= script_font * 0.4;
    }
    if fmt.subscript == Some(true) {
        y += script_font * 0.2;
    }
    y
}

fn horizontal_scale_of(fmt: &RunFormattingIn) -> Option<Number> {
    match fmt.horizontal_scale {
        Some(scale) if (scale - 100.0).abs() > f64::EPSILON => Some(px(scale)),
        _ => None,
    }
}

fn text_shadow_of(fmt: &RunFormattingIn) -> Option<String> {
    if fmt.emboss == Some(true) {
        Some("emboss".to_string())
    } else if fmt.imprint == Some(true) {
        Some("imprint".to_string())
    } else if fmt.text_shadow == Some(true) {
        Some("shadow".to_string())
    } else {
        None
    }
}

fn text_primitive_requires_browser_path(fmt: &RunFormattingIn) -> bool {
    fmt.all_caps == Some(true)
        || fmt.small_caps == Some(true)
        || fmt.hidden == Some(true)
        || fmt.text_shadow == Some(true)
        || fmt.emboss == Some(true)
        || fmt.imprint == Some(true)
        || fmt.text_outline == Some(true)
        || fmt.emphasis_mark.is_some()
        || fmt.text_effect.is_some()
}

fn structural_revision(
    info: &RevisionInfoIn,
    scope: StructuralRevisionScope,
    kind: StructuralRevisionKind,
    row_index: Option<u64>,
    col_index: Option<u64>,
) -> StructuralRevision {
    StructuralRevision {
        scope,
        author: info.author.clone().unwrap_or_default(),
        date: info.date.clone(),
        revision_id: info
            .revision_id
            .map(|id| id.to_string())
            .unwrap_or_default(),
        kind,
        row_index,
        col_index,
    }
}

fn structural_color(kind: StructuralRevisionKind) -> &'static str {
    match kind {
        StructuralRevisionKind::Ins => REVISION_INS_COLOR,
        StructuralRevisionKind::Del => REVISION_DEL_COLOR,
        StructuralRevisionKind::Merge => REVISION_MERGE_COLOR,
    }
}

fn sdt_attrs_at_depth(group: &SdtGroupIn, depth: usize) -> SdtAttrs {
    SdtAttrs {
        group_id: group.id.clone(),
        sdt_type: group.sdt_type.clone(),
        depth: Some(depth as u64),
        tag: group.tag.clone(),
        alias: group.alias.clone(),
        lock: group.lock.clone(),
        checked: group.checked,
        bound: group.bound,
        repeating_item: group.repeating_item,
    }
}

fn sdt_path_from_groups(groups: &[SdtGroupIn]) -> Vec<SdtAttrs> {
    groups
        .iter()
        .enumerate()
        .map(|(index, group)| sdt_attrs_at_depth(group, index + 1))
        .collect()
}

pub(crate) fn sdt_attrs_from_groups(groups: &[SdtGroupIn]) -> Option<SdtAttrs> {
    let depth = groups.len();
    groups.last().map(|group| sdt_attrs_at_depth(group, depth))
}

fn stamp_sdt_range(prims: &mut [Primitive], groups: &[SdtGroupIn], overwrite: bool) {
    let Some(sdt) = sdt_attrs_from_groups(groups) else {
        return;
    };
    let path = sdt_path_from_groups(groups);
    for p in prims {
        if let Some(attrs) = doc_attrs_mut(p)
            && (overwrite || attrs.sdt.is_none())
        {
            attrs.sdt = Some(sdt.clone());
            attrs.sdt_path = path.clone();
        }
    }
}

fn same_revision_burst(a: &RevisionInfoIn, b: &RevisionInfoIn) -> bool {
    a.author.as_deref().unwrap_or("") == b.author.as_deref().unwrap_or("")
        && a.date.as_deref() == b.date.as_deref()
}

fn whole_table_revision(block: &TableBlockIn) -> Option<StructuralRevision> {
    let first = block.rows.first()?;
    if let Some(shared) = first.tracked_ins.as_ref()
        && block.rows.iter().all(|row| {
            row.tracked_ins
                .as_ref()
                .is_some_and(|r| same_revision_burst(r, shared))
        })
    {
        return Some(structural_revision(
            shared,
            StructuralRevisionScope::Table,
            StructuralRevisionKind::Ins,
            None,
            None,
        ));
    }
    if let Some(shared) = first.tracked_del.as_ref()
        && block.rows.iter().all(|row| {
            row.tracked_del
                .as_ref()
                .is_some_and(|r| same_revision_burst(r, shared))
        })
    {
        return Some(structural_revision(
            shared,
            StructuralRevisionScope::Table,
            StructuralRevisionKind::Del,
            None,
            None,
        ));
    }
    None
}

fn row_structural_revision(row: &TableRowIn, row_index: usize) -> Option<StructuralRevision> {
    row.tracked_del
        .as_ref()
        .map(|info| {
            structural_revision(
                info,
                StructuralRevisionScope::Row,
                StructuralRevisionKind::Del,
                Some(row_index as u64),
                None,
            )
        })
        .or_else(|| {
            row.tracked_ins.as_ref().map(|info| {
                structural_revision(
                    info,
                    StructuralRevisionScope::Row,
                    StructuralRevisionKind::Ins,
                    Some(row_index as u64),
                    None,
                )
            })
        })
}

fn row_parent_revision_id(row: &TableRowIn) -> Option<i64> {
    row.tracked_ins
        .as_ref()
        .and_then(|r| r.revision_id)
        .or_else(|| row.tracked_del.as_ref().and_then(|r| r.revision_id))
}

/// Returns CSS font shorthand for a run.
fn css_font(fmt: &RunFormattingIn) -> String {
    let size = px(effective_font_px_of(fmt));
    let weight = if fmt.bold == Some(true) { 700 } else { 400 };
    let family = fmt.font_family.as_deref().unwrap_or(DEFAULT_FONT_FAMILY);
    let variant = if fmt.small_caps == Some(true) {
        "small-caps "
    } else {
        ""
    };
    if fmt.italic == Some(true) {
        format!("italic {variant}{weight} {size}px {family}, sans-serif")
    } else {
        format!("{variant}{weight} {size}px {family}, sans-serif")
    }
}

/// Resolves deletion, hyperlink, and explicit run colors.
fn run_color(fmt: &RunFormattingIn) -> String {
    if fmt.is_deletion == Some(true) {
        return "#c62828".to_string();
    }
    if let Some(c) = &fmt.color {
        return c.clone();
    }
    if let Some(link) = &fmt.hyperlink
        && link.no_default_style != Some(true)
    {
        return "#0563c1".to_string();
    }
    "#000000".to_string()
}

pub(crate) fn sanitized_href(href: Option<&str>) -> Option<String> {
    let raw = href?;
    let probe = raw
        .replace(['\t', '\n', '\r'], "")
        .trim_start_matches(|c: char| c <= '\u{20}')
        .to_string();
    if probe.is_empty() {
        return None;
    }
    let Some(colon) = probe.find(':') else {
        return Some(raw.to_string());
    };
    let scheme = &probe[..colon];
    let mut chars = scheme.chars();
    let Some(first) = chars.next() else {
        return Some(raw.to_string());
    };
    if !first.is_ascii_alphabetic()
        || !chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '.' | '-'))
    {
        return Some(raw.to_string());
    }
    if matches!(
        scheme.to_ascii_lowercase().as_str(),
        "http" | "https" | "mailto" | "tel" | "ftp"
    ) {
        Some(raw.to_string())
    } else {
        None
    }
}

fn hyperlink_href(fmt: &RunFormattingIn) -> Option<String> {
    sanitized_href(fmt.hyperlink.as_ref().and_then(|h| h.href.as_deref()))
}

// Strong directional character ranges for automatic paragraph direction.
fn is_rtl_strong(c: char) -> bool {
    matches!(u32::from(c),
        0x0590..=0x085F | 0x08A0..=0x08FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF)
}

fn is_ltr_strong(c: char) -> bool {
    matches!(u32::from(c),
        0x0041..=0x005A | 0x0061..=0x007A | 0x00C0..=0x02B8 | 0x0370..=0x0589
        | 0x10A0..=0x10FF | 0x1E00..=0x1FFF)
}

/// base-direction detection for paragraphs without explicit w:bidi: only
/// paragraphs carrying at least one w:rtl run are candidates; the base then
/// follows the first strong directional character (dir="auto" rule). (#719)
fn paragraph_base_is_rtl(block: &ParagraphBlockIn) -> bool {
    let mut has_rtl_run = false;
    for run in &block.runs {
        if let RunIn::Text(t) = run
            && t.fmt.rtl == Some(true)
        {
            has_rtl_run = true;
            break;
        }
    }
    if !has_rtl_run {
        return false;
    }
    for run in &block.runs {
        if let RunIn::Text(t) = run {
            for c in t.text.chars() {
                if is_rtl_strong(c) {
                    return true;
                }
                if is_ltr_strong(c) {
                    return false;
                }
            }
        }
    }
    true
}

fn is_floating_wrap_type(wrap: Option<&str>) -> bool {
    matches!(
        wrap,
        Some("square") | Some("tight") | Some("through") | Some("behind") | Some("inFront")
    )
}

/// Returns whether an image is positioned outside inline flow.
fn is_floating_image_run(run: &ImageRunIn) -> bool {
    is_floating_wrap_type(run.wrap_type.as_deref()) || run.display_mode.as_deref() == Some("float")
}

/// parse "rotate(NNdeg)" out of a CSS transform string, normalized to [0, 360)
pub(crate) fn rotation_degrees(transform: Option<&str>) -> f64 {
    let Some(t) = transform else { return 0.0 };
    let Some(idx) = t.find("rotate(") else {
        return 0.0;
    };
    let rest = &t[idx + 7..];
    let Some(end) = rest.find("deg") else {
        return 0.0;
    };
    let deg: f64 = rest[..end].trim().parse().unwrap_or(0.0);
    ((deg % 360.0) + 360.0) % 360.0
}

/// hard cap on image alt text — `wp:docPr descr` is attacker-controlled file
/// data, so an unbounded value must never ride into every consumer's copy of
/// the display list (a11y mirror DOM, canvas hosts, serialized snapshots)
pub const MAX_ALT_TEXT_CHARS: usize = 2048;

/// Alt text for an image primitive: an empty value is dropped rather than
/// emitted, and an oversized one truncates on a character boundary.
pub(crate) fn capped_alt_text(alt: Option<&str>) -> Option<String> {
    let alt = alt?;
    if alt.is_empty() {
        return None;
    }
    Some(alt.chars().take(MAX_ALT_TEXT_CHARS).collect())
}

fn emit_watermark(prims: &mut Vec<Primitive>, watermark: &WatermarkIn, page: &PageIn) {
    match watermark {
        WatermarkIn::Text(wm) => emit_text_watermark(prims, wm, page),
        WatermarkIn::Picture(wm) => emit_picture_watermark(prims, wm, page),
    }
}

fn watermark_auto_font_px(text: &str, available_width_px: f64) -> f64 {
    let chars = text.trim().chars().count().max(1) as f64;
    let size = available_width_px / (chars * 0.62);
    size.clamp(24.0, 180.0)
}

fn emit_text_watermark(prims: &mut Vec<Primitive>, wm: &TextWatermarkIn, page: &PageIn) {
    if wm.text.is_empty() {
        return;
    }
    let text: String = wm.text.chars().take(MAX_ALT_TEXT_CHARS).collect();
    let content_width = page.size.w - page.margins.left - page.margins.right;
    let target_width = if wm.layout.as_deref() == Some("diagonal") {
        content_width * 1.3
    } else {
        content_width
    };
    let font_px = wm
        .font_size
        .map(|pt| pt * 96.0 / 72.0)
        .filter(|px| px.is_finite() && *px > 0.0)
        .unwrap_or_else(|| watermark_auto_font_px(&text, target_width));
    let width = (text.trim().chars().count().max(1) as f64) * font_px * 0.62;
    let x = (page.size.w - width) / 2.0;
    // textRunRect centers at baseline - 0.2em; choose the baseline so the
    // primitive's geometry center is exactly the page center before rotation.
    let baseline = page.size.h / 2.0 + font_px * 0.2;
    let family = wm.font.as_deref().unwrap_or(DEFAULT_FONT_FAMILY);
    let font = format!("700 {}px {}, sans-serif", px(font_px), family);
    let opacity = if wm.semitransparent == Some(true) {
        0.5
    } else {
        0.85
    };
    let rotation = if wm.layout.as_deref() == Some("diagonal") {
        Some(px(-45.0))
    } else {
        None
    };

    let attrs = DocAttrs {
        decorative: Some(wm.decorative.unwrap_or(true)),
        ..DocAttrs::default()
    };
    prims.push(Primitive::Text(TextRunPrimitive {
        text,
        x: px(x),
        baseline_y: px(baseline),
        width: px(width),
        paint_clip: None,
        font,
        color: wm.color.clone().unwrap_or_else(|| "#C0C0C0".to_string()),
        letter_spacing: None,
        word_spacing: None,
        rtl: None,
        opacity: Some(px(opacity)),
        rotation_deg: rotation,
        horizontal_scale: None,
        all_caps: false,
        small_caps: false,
        hidden: false,
        text_shadow: None,
        text_outline: false,
        emphasis_mark: None,
        text_effect: None,
        attrs,
    }));
}

fn emit_picture_watermark(prims: &mut Vec<Primitive>, wm: &PictureWatermarkIn, page: &PageIn) {
    let rel_id = wm
        .data_url
        .as_deref()
        .or(wm.rel_id.as_deref())
        .unwrap_or("");
    if rel_id.is_empty() {
        return;
    }
    let content_width = page.size.w - page.margins.left - page.margins.right;
    let natural_width = wm
        .width_emu
        .map(emu_to_px)
        .filter(|w| w.is_finite() && *w > 0.0)
        .unwrap_or(content_width * 0.75);
    let natural_height = wm
        .height_emu
        .map(emu_to_px)
        .filter(|h| h.is_finite() && *h > 0.0)
        .unwrap_or(natural_width);
    let scale = wm
        .scale
        .filter(|s| s.is_finite() && *s > 0.0)
        .unwrap_or(1.0);
    let w = natural_width * scale;
    let h = natural_height * scale;
    let washout = wm.washout == Some(true);

    prims.push(Primitive::Image(ImagePrimitive {
        rel_id: rel_id.to_string(),
        x: px((page.size.w - w) / 2.0),
        y: px((page.size.h - h) / 2.0),
        w: px(w),
        h: px(h),
        rotation_deg: None,
        opacity: if washout { Some(px(0.5)) } else { None },
        filter: if washout {
            Some("brightness(1.4) contrast(0.4)".to_string())
        } else {
            None
        },
        decorative: wm.decorative.unwrap_or(true),
        crop: None,
        alt_text: None,
        attrs: DocAttrs::default(),
    }));
}

fn crop_of(run: &ImageRunIn) -> Option<Crop> {
    let t = run.crop_top.unwrap_or(0.0);
    let r = run.crop_right.unwrap_or(0.0);
    let b = run.crop_bottom.unwrap_or(0.0);
    let l = run.crop_left.unwrap_or(0.0);
    if t == 0.0 && r == 0.0 && b == 0.0 && l == 0.0 {
        return None;
    }
    Some(Crop {
        top: px(t),
        right: px(r),
        bottom: px(b),
        left: px(l),
    })
}

fn crop_of_block(block: &ImageBlockIn) -> Option<Crop> {
    let crop = block.crop.as_ref()?;
    Some(Crop {
        top: px(crop.top.unwrap_or(0.0)),
        right: px(crop.right.unwrap_or(0.0)),
        bottom: px(crop.bottom.unwrap_or(0.0)),
        left: px(crop.left.unwrap_or(0.0)),
    })
}

fn transform_has_flip(transform: Option<&str>, axis: char) -> bool {
    let needle = if axis == 'x' {
        "scaleX(-1)"
    } else {
        "scaleY(-1)"
    };
    transform.is_some_and(|value| value.replace(' ', "").contains(needle))
}

fn content_frame(
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    bounds: Option<&RotationBoundsIn>,
) -> Option<ContentFrame> {
    let bounds = bounds?;
    Some(ContentFrame {
        x: Some(px(x + bounds.offset_x.unwrap_or(0.0))),
        y: Some(px(y + bounds.offset_y.unwrap_or(0.0))),
        w: Some(px(width)),
        h: Some(px(height)),
    })
}

fn stamp_hyperlink_attrs(
    attrs: &mut DocAttrs,
    link: Option<&HyperlinkIn>,
    legacy_href: Option<&str>,
) {
    let href = link.and_then(|value| value.href.as_deref()).or(legacy_href);
    attrs.href = sanitized_href(href);
    if let Some(link) = link {
        attrs.tooltip = link.tooltip.clone();
        attrs.link_title = link.tooltip.clone();
        attrs.link_target = link.target.clone();
        attrs.link_history = link.history;
        attrs.link_doc_location = link.doc_location.clone();
    }
}

fn stamp_image_run_attrs(attrs: &mut DocAttrs, run: &ImageRunIn, x: f64, y: f64) {
    stamp_hyperlink_attrs(attrs, run.hyperlink.as_ref(), run.hlink_href.as_deref());
    attrs.image_flip_h = (run.flip_h == Some(true)
        || transform_has_flip(run.transform.as_deref(), 'x'))
    .then_some(true);
    attrs.image_flip_v = (run.flip_v == Some(true)
        || transform_has_flip(run.transform.as_deref(), 'y'))
    .then_some(true);
    attrs.content_frame = content_frame(x, y, run.width, run.height, run.rotation_bounds.as_ref());
    attrs.effects = run.effects.clone();
    attrs.border = run.outline.clone();
    if run.is_insertion == Some(true) || run.is_deletion == Some(true) {
        attrs.revision = Some(Revision {
            author: run.change_author.clone().unwrap_or_default(),
            date: run.change_date.clone().unwrap_or_default(),
            revision_id: run
                .change_revision_id
                .map(|id| id.to_string())
                .unwrap_or_default(),
            kind: if run.is_insertion == Some(true) {
                RevisionKind::Ins
            } else {
                RevisionKind::Del
            },
        });
    }
}

fn stamp_image_block_attrs(attrs: &mut DocAttrs, block: &ImageBlockIn, x: f64, y: f64) {
    attrs.href = sanitized_href(block.hlink_href.as_deref());
    attrs.link_title = block.hlink_title.clone();
    attrs.tooltip = block.hlink_title.clone();
    attrs.image_flip_h = (block.flip_h == Some(true)
        || transform_has_flip(block.transform.as_deref(), 'x'))
    .then_some(true);
    attrs.image_flip_v = (block.flip_v == Some(true)
        || transform_has_flip(block.transform.as_deref(), 'y'))
    .then_some(true);
    attrs.content_frame = content_frame(
        x,
        y,
        block.width,
        block.height,
        block.rotation_bounds.as_ref(),
    );
    attrs.effects = block.effects.clone();
    attrs.border = block.outline.clone();
}

fn image_layout_width(run: &ImageRunIn) -> f64 {
    run.rotation_bounds
        .as_ref()
        .and_then(|bounds| bounds.width)
        .unwrap_or(run.width)
}

fn image_layout_height(run: &ImageRunIn) -> f64 {
    run.rotation_bounds
        .as_ref()
        .and_then(|bounds| bounds.height)
        .unwrap_or(run.height)
}

/// one run's visible slice on a laid-out line: the (possibly sliced) run plus
/// its on-line text; non-text runs pass through whole with empty text
struct ResolvedSegment<'a> {
    run: &'a RunIn,
    /// for boundary text runs: the sliced text + shifted document positions
    text: String,
    pm_start: Option<i64>,
    pm_end: Option<i64>,
}

fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

/// Slice using UTF-16 offsets. Start/end are snapped to
/// scalar boundaries so malformed hand-authored envelopes cannot split a
/// surrogate pair; authoritative cluster metadata already lands on grapheme
/// boundaries and therefore passes through unchanged.
fn slice_utf16(text: &str, start: usize, end: usize) -> String {
    slice_utf16_with_total(text, utf16_len(text), start, end)
}

/// [`slice_utf16`] with the text's utf16 length precomputed, so per-line run
/// slicing measures each run once instead of once per bound.
fn slice_utf16_with_total(text: &str, total: usize, start: usize, end: usize) -> String {
    let start = start.min(total);
    let end = end.max(start).min(total);
    let mut units = 0usize;
    text.chars()
        .filter_map(|ch| {
            let next = units + ch.len_utf16();
            let keep = units >= start && next <= end;
            units = next;
            keep.then_some(ch)
        })
        .collect()
}

fn resolve_line_segments<'a>(runs: &'a [RunIn], line: &LineIn) -> Vec<ResolvedSegment<'a>> {
    let mut out = Vec::new();
    for run_index in line.head_run..=line.tail_run {
        let Some(run) = runs.get(run_index) else {
            continue;
        };
        match run {
            RunIn::Text(t) => {
                let total = utf16_len(&t.text);
                let start = if run_index == line.head_run {
                    line.head_char.min(total)
                } else {
                    0
                };
                let end = if run_index == line.tail_run {
                    line.tail_char.min(total)
                } else {
                    total
                };
                let end = end.max(start);
                let text = slice_utf16_with_total(&t.text, total, start, end);
                // an unsliced run keeps its own document span; a boundary slice
                // shifts the positions to match (text runs are 1 position per char)
                let (pm_start, pm_end) = if start == 0 && end == total {
                    (t.pm_start, t.pm_end.or(t.pm_start.map(|p| p + end as i64)))
                } else {
                    match t.pm_start {
                        Some(p) => (Some(p + start as i64), Some(p + end as i64)),
                        None => (None, None),
                    }
                };
                out.push(ResolvedSegment {
                    run,
                    text,
                    pm_start,
                    pm_end,
                });
            }
            RunIn::Tab(t) => out.push(ResolvedSegment {
                run,
                text: String::new(),
                pm_start: t.pm_start,
                pm_end: t.pm_end,
            }),
            RunIn::Image(t) => out.push(ResolvedSegment {
                run,
                text: String::new(),
                pm_start: t.pm_start,
                pm_end: t.pm_end,
            }),
            RunIn::LineBreak(t) => out.push(ResolvedSegment {
                run,
                text: String::new(),
                pm_start: t.pm_start,
                pm_end: t.pm_start,
            }),
            RunIn::Field(t) => out.push(ResolvedSegment {
                run,
                text: String::new(),
                pm_start: t.pm_start,
                pm_end: t.pm_end,
            }),
            RunIn::Unsupported => {}
        }
    }
    out
}

struct LineTextItem<'a> {
    text: String,
    fmt: &'a RunFormattingIn,
    pm_start: Option<i64>,
    pm_end: Option<i64>,
    width: f64,
    level: u8,
    source_start: usize,
    source_end: usize,
    logical_order: Option<u64>,
    exact_advance: bool,
    /// source field run when this item paints a field result — carries the
    /// inert a11y identity (type/instruction) onto the emitted primitives
    field: Option<&'a FieldRunIn>,
}

enum LinePaintItem<'a> {
    Text(LineTextItem<'a>),
    Tab {
        run: &'a TabRunIn,
        width: f64,
        level: u8,
        logical_order: Option<u64>,
    },
    Image {
        run: &'a ImageRunIn,
        pm_start: Option<i64>,
        pm_end: Option<i64>,
        single_image_line: bool,
        level: u8,
        logical_order: Option<u64>,
    },
    LineBreak {
        pm_start: Option<i64>,
        level: u8,
    },
}

impl LinePaintItem<'_> {
    fn level(&self) -> u8 {
        match self {
            LinePaintItem::Text(item) => item.level,
            LinePaintItem::Tab { level, .. }
            | LinePaintItem::Image { level, .. }
            | LinePaintItem::LineBreak { level, .. } => *level,
        }
    }

    fn width(&self) -> f64 {
        match self {
            LinePaintItem::Text(item) => item.width,
            LinePaintItem::Tab { width, .. } => *width,
            LinePaintItem::Image { run, .. } => run
                .rotation_bounds
                .as_ref()
                .and_then(|bounds| bounds.width)
                .unwrap_or(run.width),
            LinePaintItem::LineBreak { .. } => 0.0,
        }
    }
}

fn base_bidi_direction(is_rtl: bool) -> ooxml_text::BaseDirection {
    if is_rtl {
        ooxml_text::BaseDirection::Rtl
    } else {
        ooxml_text::BaseDirection::Ltr
    }
}

fn base_bidi_level(is_rtl: bool) -> u8 {
    if is_rtl { 1 } else { 0 }
}

fn shape_direction_for_level(level: u8) -> ooxml_text::ShapeDirection {
    if ooxml_text::level_is_rtl(level) {
        ooxml_text::ShapeDirection::Rtl
    } else {
        ooxml_text::ShapeDirection::Ltr
    }
}

fn bidi_char_levels(text: &str, base: ooxml_text::BaseDirection) -> Vec<u8> {
    if text.is_empty() {
        return Vec::new();
    }
    let paras = ooxml_text::bidi_paragraphs(text, base);
    let mut runs = paras.iter().flat_map(|p| p.runs.iter());
    let mut cur = runs.next();
    let mut levels = Vec::new();
    for (byte, _) in text.char_indices() {
        while let Some(r) = cur {
            if byte < r.end {
                break;
            }
            cur = runs.next();
        }
        levels.push(cur.map_or_else(
            || base_bidi_level(base == ooxml_text::BaseDirection::Rtl),
            |r| r.level,
        ));
    }
    levels
}

#[allow(clippy::too_many_arguments)]
fn push_bidi_text_items<'a>(
    out: &mut Vec<LinePaintItem<'a>>,
    text: &str,
    fmt: &'a RunFormattingIn,
    pm_start: Option<i64>,
    pm_end: Option<i64>,
    total_width: f64,
    levels: &[u8],
    level_cursor: &mut usize,
    default_level: u8,
    word_space_extra: f64,
    field: Option<&'a FieldRunIn>,
) {
    let chars: Vec<char> = text.chars().collect();
    if chars.is_empty() {
        out.push(LinePaintItem::Text(LineTextItem {
            text: String::new(),
            fmt,
            pm_start,
            pm_end,
            width: total_width,
            level: default_level,
            source_start: 0,
            source_end: 0,
            logical_order: fmt.logical_order,
            exact_advance: false,
            field,
        }));
        return;
    }

    let total_chars = chars.len();
    let total_scalars = fallback_scalar_count(text, fmt.all_caps);
    let mut utf16_offsets = Vec::with_capacity(total_chars + 1);
    utf16_offsets.push(0usize);
    let mut utf16_units = 0usize;
    for character in &chars {
        utf16_units += character.len_utf16();
        utf16_offsets.push(utf16_units);
    }
    let mut start = 0usize;
    while start < total_chars {
        let level = levels
            .get(*level_cursor + start)
            .copied()
            .unwrap_or(default_level);
        let mut end = start + 1;
        while end < total_chars
            && levels
                .get(*level_cursor + end)
                .copied()
                .unwrap_or(default_level)
                == level
        {
            end += 1;
        }

        let slice: String = chars[start..end].iter().collect();
        let slice_scalars = fallback_scalar_count(&slice, fmt.all_caps);
        let spaces = slice.chars().filter(|&ch| ch == ' ').count();
        let base_width = total_width * slice_scalars as f64 / total_scalars as f64;
        let width = if word_space_extra > 0.0 {
            // `total_width` already includes the stretched spaces. Remove the
            // segment-level equal-share stretch and reapply it per bidi slice.
            let total_spaces = text.chars().filter(|&ch| ch == ' ').count();
            let unstretched = total_width - word_space_extra * total_spaces as f64;
            (unstretched * slice_scalars as f64 / total_scalars as f64)
                + word_space_extra * spaces as f64
        } else {
            base_width
        };

        let slice_pm_end = if end == total_chars {
            pm_end.or_else(|| pm_start.map(|p| p + utf16_offsets[end] as i64))
        } else {
            pm_start.map(|p| p + utf16_offsets[end] as i64)
        };

        out.push(LinePaintItem::Text(LineTextItem {
            text: slice,
            fmt,
            pm_start: pm_start.map(|p| p + utf16_offsets[start] as i64),
            pm_end: slice_pm_end,
            width,
            level,
            source_start: utf16_offsets[start],
            source_end: utf16_offsets[end],
            logical_order: fmt.logical_order,
            exact_advance: false,
            field,
        }));
        start = end;
    }
    *level_cursor += total_chars;
}

// ---------------------------------------------------------------------------
// builder
// ---------------------------------------------------------------------------

/// Per-page resolved widths for one PAGE/NUMPAGES field run. `fallback` is the
/// width the measure baked into the line (from the field's fallback text);
/// `per_page[i]` is the width of its resolved text on layout page `i`.
pub(crate) struct FieldWidthEntry {
    pub(crate) fallback: f64,
    pub(crate) per_page: Vec<f64>,
}

/// field document position → per-page widths (see [`FieldWidthEntry`]).
pub(crate) type FieldWidthMap = HashMap<i64, FieldWidthEntry>;

/// Everything emission needs that varies per page or per band: the numbers
/// PAGE and NUMPAGES resolve to, the shaping fonts, and the field widths scoped
/// to the header/footer variant being composed.
pub(crate) struct RenderCtx<'a> {
    pub(crate) page_number: u64,
    /// Zero-based key into per-page field-width arrays.
    /// Distinct from `page_number`, which restarts per section.
    pub(crate) page_index: usize,
    pub(crate) total_pages: u64,
    /// Shaping fonts for glyph-run emission.
    pub(crate) shape: Option<&'a ShapeFonts<'a>>,
    /// Per-page PAGE/NUMPAGES widths for header/footer field lines.
    pub(crate) field_widths: Option<&'a FieldWidthMap>,
}

impl RenderCtx<'_> {
    /// `(fallback_width, resolved_width_on_this_page)` for a field run, when the
    /// input supplied a per-page width keyed on its document position.
    fn field_width(&self, pm_start: Option<i64>) -> Option<(f64, f64)> {
        let entry = self.field_widths?.get(&pm_start?)?;
        let resolved = entry.per_page.get(self.page_index).copied()?;
        Some((entry.fallback, resolved))
    }
}

/// The measurement font store plus the input's `fontChains` (u32 ids converted
/// to `FontId` once), the two things `emit_text_segment` needs to shape a run
/// into a [`GlyphRunPrimitive`]. Built once per `build_display_list` and
/// borrowed by every page/HF `RenderCtx`.
pub(crate) struct ShapeFonts<'a> {
    store: &'a ooxml_text::FontStore,
    chains: HashMap<String, Vec<ooxml_text::FontId>>,
}

impl<'a> ShapeFonts<'a> {
    /// Build the shaping context from the input's `fontChains`, or `None` when
    /// the input carries no chains (the browser-measured path). The u32 ids are
    /// converted to `FontId` here so lookups downstream are cheap; ids that
    /// don't belong to `store` fail at shape time and route their run back to
    /// the `TextRunPrimitive` fallback.
    fn build(input: &BuildInput, store: &'a ooxml_text::FontStore) -> Option<ShapeFonts<'a>> {
        if input.font_chains.is_empty() {
            return None;
        }
        let chains = input
            .font_chains
            .iter()
            .map(|(k, ids)| {
                (
                    k.clone(),
                    ids.iter()
                        .map(|&id| ooxml_text::FontId::from_u32(id))
                        .collect(),
                )
            })
            .collect();
        Some(ShapeFonts { store, chains })
    }

    /// Fallback chain for a `(family, bold, italic)` combination, keyed exactly
    /// like the measure input (`"<family lowercase>|<b>|<i>"`). `None` when the
    /// run's family has no chain — that run falls back to `TextRunPrimitive`.
    fn chain_for(&self, family: &str, bold: bool, italic: bool) -> Option<&[ooxml_text::FontId]> {
        let key = format!(
            "{}|{}|{}",
            family.to_lowercase(),
            u8::from(bold),
            u8::from(italic)
        );
        self.chains.get(&key).map(|v| v.as_slice())
    }

    /// Resolve `ch` to the first covering font in `chain`, else the chain's
    /// terminal font — the same policy as ooxml-text `prepare::resolve_with_fallback`
    /// (the host guarantees the chain ends in a broad-coverage last-resort face,
    /// so an uncovered char shapes as that face's `.notdef` box rather than
    /// routing the whole run to the browser).
    fn resolve(&self, chain: &[ooxml_text::FontId], ch: char) -> Option<ooxml_text::FontId> {
        self.store
            .resolve(chain, ch)
            .or_else(|| chain.last().copied())
    }
}

fn emit_column_separators(prims: &mut Vec<Primitive>, page: &PageIn) {
    let Some(columns) = &page.columns else { return };
    if columns.separator != Some(true) || columns.count <= 1 {
        return;
    }
    let content_width = (page.size.w - page.margins.left - page.margins.right).max(0.0);
    let content_bottom = page.size.h - page.margins.bottom;
    if columns.equal_width == Some(false) && !columns.columns.is_empty() {
        let mut cursor = page.margins.left;
        for index in 0..columns.count.saturating_sub(1) {
            let width = columns
                .columns
                .get(index)
                .and_then(|column| column.width)
                .unwrap_or(0.0);
            let space = columns
                .columns
                .get(index)
                .and_then(|column| column.space)
                .unwrap_or(columns.gap);
            cursor += width;
            prims.push(Primitive::Line(LinePrimitive {
                x1: px(cursor + space / 2.0),
                y1: px(page.margins.top),
                x2: px(cursor + space / 2.0),
                y2: px(content_bottom),
                stroke_width: px(0.5),
                color: "#000000".to_string(),
                dash: None,
                role: Some(LineRole::Separator),
                border_style: Some(DisplayBorderStyle::Solid),
                secondary_color: None,
                opacity: None,
                border_owner: None,
                attrs: DocAttrs::default(),
            }));
            cursor += space;
        }
        return;
    }

    let width = (content_width - (columns.count - 1) as f64 * columns.gap) / columns.count as f64;
    for index in 0..columns.count - 1 {
        let x = page.margins.left
            + (index + 1) as f64 * width
            + index as f64 * columns.gap
            + columns.gap / 2.0;
        prims.push(Primitive::Line(LinePrimitive {
            x1: px(x),
            y1: px(page.margins.top),
            x2: px(x),
            y2: px(content_bottom),
            stroke_width: px(0.5),
            color: "#000000".to_string(),
            dash: None,
            role: Some(LineRole::Separator),
            border_style: Some(DisplayBorderStyle::Solid),
            secondary_color: None,
            opacity: None,
            border_owner: None,
            attrs: DocAttrs::default(),
        }));
    }
}

/// Exact body content/column boxes for interaction queries.
fn page_content_geometry(page: &PageIn) -> (DisplayBounds, Vec<DisplayBounds>) {
    let content_width = (page.size.w - page.margins.left - page.margins.right).max(0.0);
    let content_height = (page.size.h - page.margins.top - page.margins.bottom).max(0.0);
    let content = DisplayBounds {
        x: px(page.margins.left),
        y: px(page.margins.top),
        width: px(content_width),
        height: px(content_height),
    };

    let Some(columns) = &page.columns else {
        return (content.clone(), vec![content]);
    };
    let count = columns.count.clamp(1, 45);
    if columns.equal_width == Some(false) && !columns.columns.is_empty() {
        let mut widths = Vec::with_capacity(count);
        let mut gaps = Vec::with_capacity(count);
        for index in 0..count {
            let width = columns
                .columns
                .get(index)
                .and_then(|column| column.width)
                .filter(|value| value.is_finite() && *value > 0.0)
                .unwrap_or(0.0);
            let gap = columns
                .columns
                .get(index)
                .and_then(|column| column.space)
                .filter(|value| value.is_finite() && *value >= 0.0)
                .unwrap_or_else(|| columns.gap.max(0.0));
            widths.push(width);
            gaps.push(gap);
        }
        let gap_total: f64 = gaps.iter().take(count.saturating_sub(1)).sum();
        let known_total: f64 = widths.iter().sum();
        let missing = widths.iter().filter(|width| **width <= 0.0).count();
        let fallback = if missing > 0 {
            (content_width - gap_total - known_total).max(0.0) / missing as f64
        } else {
            0.0
        };
        for width in &mut widths {
            if *width <= 0.0 {
                *width = fallback;
            }
        }
        let widths_total: f64 = widths.iter().sum();
        if widths_total + gap_total > content_width && widths_total > 0.0 {
            let scale = (content_width - gap_total).max(0.0) / widths_total.max(1.0);
            for width in &mut widths {
                *width *= scale;
            }
        }
        let mut x = page.margins.left;
        let mut bounds = Vec::with_capacity(count);
        for index in 0..count {
            bounds.push(DisplayBounds {
                x: px(x),
                y: px(page.margins.top),
                width: px(widths[index]),
                height: px(content_height),
            });
            x += widths[index] + if index + 1 < count { gaps[index] } else { 0.0 };
        }
        return (content, bounds);
    }

    let gap = if columns.gap.is_finite() {
        columns.gap.max(0.0)
    } else {
        0.0
    };
    let width = ((content_width - count.saturating_sub(1) as f64 * gap) / count as f64).max(0.0);
    let bounds = (0..count)
        .map(|index| DisplayBounds {
            x: px(page.margins.left + index as f64 * (width + gap)),
            y: px(page.margins.top),
            width: px(width),
            height: px(content_height),
        })
        .collect();
    (content, bounds)
}

const NOTE_COLUMN_GAP_PX: f64 = 24.0;
const NOTE_SEPARATOR_HEIGHT_PX: f64 = 12.0;
/// note reference-label cap (display numbers / custom marks are tiny)
pub const MAX_NOTE_LABEL_CHARS: usize = 64;

fn note_partitions<'a>(notes: &'a [NoteItemIn], columns: usize) -> Vec<Vec<&'a NoteItemIn>> {
    let columns = columns.max(1);
    if columns == 1 || notes.len() <= 1 {
        return vec![notes.iter().collect()];
    }
    let total = notes
        .iter()
        .map(|note| note.height.unwrap_or(0.0))
        .sum::<f64>();
    let target = total / columns as f64;
    let mut out: Vec<Vec<&NoteItemIn>> = vec![Vec::new()];
    let mut height = 0.0;
    for note in notes {
        let note_height = note.height.unwrap_or(0.0);
        if out.len() < columns && height > 0.0 && height + note_height / 2.0 > target {
            out.push(Vec::new());
            height = 0.0;
        }
        if let Some(column) = out.last_mut() {
            column.push(note);
        }
        height += note_height;
    }
    out
}

/// Paint-group id shared by every primitive of one note. Grouping a note
/// area's primitives back into stories means rebuilding it, which hit testing
/// and the a11y mirror both do.
pub(crate) fn note_group_id(kind: &str, id: i64) -> String {
    format!("{kind}-{id}")
}

/// Marks a note's primitives with its paint group. Their document positions
/// stay the note story's own, so a point or range inside a note addresses that
/// story; the body-reference range is linkage metadata and rides on the
/// region's `notes` entry instead.
fn stamp_note_item(prims: &mut [Primitive], start: usize, note: &NoteItemIn, kind: &str) {
    let Some(id) = note.id else { return };
    let group_id = note_group_id(kind, id);
    for primitive in &mut prims[start..] {
        if let Some(attrs) = doc_attrs_mut(primitive) {
            attrs.group_id = Some(group_id.clone());
        }
    }
}

fn emit_note_item(
    prims: &mut Vec<Primitive>,
    note: &NoteItemIn,
    kind: &str,
    x: f64,
    y: f64,
    width: f64,
    ctx: &RenderCtx<'_>,
) -> f64 {
    let start = prims.len();
    let mut cursor = 0.0;
    for (block, measure) in note.blocks.iter().zip(&note.measures) {
        match (block, measure) {
            (BlockIn::Paragraph(block), MeasureIn::Paragraph(measure)) => {
                let before = block
                    .attrs
                    .as_ref()
                    .and_then(|attrs| attrs.spacing)
                    .and_then(|spacing| spacing.before)
                    .unwrap_or(0.0);
                cursor += before;
                let fragment = ParagraphFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y: y + cursor,
                    width,
                    height: measure.total_height,
                    from_line: 0,
                    to_line: measure.lines.len(),
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_paragraph_fragment(
                    prims,
                    &fragment,
                    block,
                    measure,
                    ctx,
                    x,
                    y + cursor,
                    None,
                    None,
                    true,
                    true,
                );
                cursor += measure.total_height;
            }
            (BlockIn::Table(block), MeasureIn::Table(measure)) => {
                let fragment = TableFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y: y + cursor,
                    width: measure.total_width,
                    height: measure.total_height,
                    row_start: 0,
                    row_end: block.rows.len(),
                    clip_top: None,
                    clip_bottom: None,
                    header_row_count: None,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_table_fragment(prims, &fragment, block, measure, ctx);
                cursor += measure.total_height;
            }
            (BlockIn::Image(block), MeasureIn::Image(measure)) => {
                let mut attrs = BlockRef::of(&block.id).attrs();
                attrs.doc_start = block.pm_start;
                attrs.doc_end = block.pm_end;
                stamp_image_block_attrs(&mut attrs, block, x, y + cursor);
                let rotation = block
                    .rotation_deg
                    .unwrap_or_else(|| rotation_degrees(block.transform.as_deref()));
                prims.push(Primitive::Image(ImagePrimitive {
                    rel_id: block.src.clone(),
                    x: px(x),
                    y: px(y + cursor),
                    w: px(measure.width),
                    h: px(measure.height),
                    rotation_deg: (rotation != 0.0).then(|| px(rotation)),
                    opacity: block.opacity.map(px),
                    filter: None,
                    decorative: block.decorative.unwrap_or(false),
                    crop: crop_of_block(block),
                    alt_text: capped_alt_text(block.alt.as_deref()),
                    attrs,
                }));
                cursor += measure.height;
            }
            (BlockIn::TextBox(block), MeasureIn::TextBox(measure)) => {
                let fragment = TextBoxFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y: y + cursor,
                    width,
                    height: measure.height,
                    pm_start: None,
                    pm_end: None,
                    is_floating: None,
                    z_index: None,
                };
                emit_text_box_fragment(prims, &fragment, block, measure, ctx);
                cursor += measure.height;
            }
            (BlockIn::Shape(block), MeasureIn::Shape(measure)) => {
                let fragment = ShapeFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y: y + cursor,
                    width: measure.width,
                    height: measure.height,
                    doc_start: block.doc_start,
                    doc_end: block.doc_end,
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                    is_anchored: None,
                    z_index: None,
                };
                emit_shape_fragment(prims, &fragment, block, ctx);
                cursor += measure.height;
            }
            (BlockIn::Chart(block), MeasureIn::Chart(measure)) => {
                let fragment = ChartFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y: y + cursor,
                    width: measure.width,
                    height: measure.height,
                    doc_start: block.doc_start,
                    doc_end: block.doc_end,
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                };
                emit_chart_fragment(prims, &fragment, block);
                cursor += measure.height;
            }
            _ => {}
        }
    }
    stamp_note_item(prims, start, note, kind);
    note.height.unwrap_or(cursor).max(cursor)
}

fn emit_note_regions(page: &PageIn, ctx: &RenderCtx<'_>) -> Vec<NoteRegion> {
    let content_width = (page.size.w - page.margins.left - page.margins.right).max(1.0);
    let mut regions = Vec::with_capacity(page.note_areas.len());
    for area in &page.note_areas {
        let kind = area.kind.as_deref().unwrap_or("footnote");
        let y = area
            .y
            .unwrap_or(page.size.h - page.margins.bottom - area.height.unwrap_or(0.0));
        let columns = area.columns.unwrap_or(1).max(1) as usize;
        let column_width =
            ((content_width - (columns - 1) as f64 * NOTE_COLUMN_GAP_PX) / columns as f64).max(1.0);
        let mut separator_primitives = Vec::new();
        if let Some(separator) = &area.separator {
            emit_note_item(
                &mut separator_primitives,
                separator,
                kind,
                page.margins.left,
                y,
                content_width,
                ctx,
            );
        } else {
            separator_primitives.push(Primitive::Line(LinePrimitive {
                x1: px(page.margins.left),
                y1: px(y),
                x2: px(page.margins.left + content_width / 3.0),
                y2: px(y),
                stroke_width: px(1.0),
                color: "#000000".to_string(),
                dash: None,
                role: Some(LineRole::Separator),
                border_style: Some(DisplayBorderStyle::Solid),
                secondary_color: None,
                opacity: None,
                border_owner: None,
                attrs: DocAttrs::default(),
            }));
        }

        let mut primitives = Vec::new();
        let partitions = note_partitions(&area.notes, columns);
        for (column, notes) in partitions.into_iter().enumerate() {
            let x = page.margins.left + column as f64 * (column_width + NOTE_COLUMN_GAP_PX);
            let mut cursor = y + NOTE_SEPARATOR_HEIGHT_PX;
            for note in notes {
                cursor += emit_note_item(&mut primitives, note, kind, x, cursor, column_width, ctx);
            }
        }
        regions.push(NoteRegion {
            kind: area.kind.clone(),
            section_id: area.section_id.clone(),
            y: area.y.map(px),
            height: area.height.map(px),
            columns: area.columns,
            separator_primitives,
            primitives,
            note_ids: area.notes.iter().filter_map(|note| note.id).collect(),
            // Only notes carrying anchor or label data get backlink metadata.
            notes: area
                .notes
                .iter()
                .filter(|note| {
                    note.anchor_doc_start.is_some()
                        || note.anchor_doc_end.is_some()
                        || note.display_label.is_some()
                })
                .map(|note| NoteRegionNote {
                    id: note.id,
                    anchor_doc_start: note.anchor_doc_start,
                    anchor_doc_end: note.anchor_doc_end,
                    label: capped_string(note.display_label.as_deref(), MAX_NOTE_LABEL_CHARS),
                })
                .collect(),
        });
    }
    regions
}

fn measured_block_height(measured: &MeasuredBlockIn) -> f64 {
    match &measured.measure {
        MeasureIn::Paragraph(measure) => measure.total_height,
        MeasureIn::Table(measure) => measure.total_height,
        MeasureIn::Image(measure) => measure.height,
        MeasureIn::TextBox(measure) => measure.height,
        MeasureIn::Shape(measure) | MeasureIn::Chart(measure) => measure.height,
        MeasureIn::Unsupported => 0.0,
    }
}

fn resolve_hf_box_position(
    position: Option<&AnchorPosIn>,
    css_float: Option<&str>,
    width: f64,
    height: f64,
    paragraph_y: f64,
    geom: &PageFloatGeom,
) -> (f64, f64) {
    let x = match position.and_then(|position| position.horizontal.as_ref()) {
        None if css_float == Some("right") => geom.content_width - width,
        None => 0.0,
        Some(axis) => {
            let band = horizontal_anchor_band(axis.relative_to.as_deref(), geom);
            match axis.align.as_deref() {
                Some("right") => band.base + band.size - width,
                Some("center") => band.base + (band.size - width) / 2.0,
                Some("left") => band.base,
                _ => band.base + axis.pos_offset.map(emu_to_px).unwrap_or(0.0),
            }
        }
    };
    let y = match position.and_then(|position| position.vertical.as_ref()) {
        None => paragraph_y,
        Some(axis) => {
            let band = vertical_anchor_band(axis.relative_to.as_deref(), paragraph_y, geom);
            match axis.align.as_deref() {
                Some("bottom") if band.size != 0.0 => band.base + band.size - height,
                Some("center") if band.size != 0.0 => band.base + (band.size - height) / 2.0,
                Some("top") => band.base,
                _ => band.base + axis.pos_offset.map(emu_to_px).unwrap_or(0.0),
            }
        }
    };
    (geom.margin_left + x, geom.margin_top + y)
}

fn recompose_hf_region(
    region: &mut HfRegion,
    hf: &HeadersFootersContentIn,
    page: &PageIn,
    page_index: usize,
    total_pages: u64,
    shape: Option<&ShapeFonts<'_>>,
) {
    let Some(variant) = hf
        .variants
        .iter()
        .rev()
        .find(|variant| variant.r_id == region.r_id && variant.kind == region.kind)
    else {
        return;
    };
    let field_widths: FieldWidthMap = variant
        .field_widths
        .iter()
        .map(|field| {
            (
                field.pm_start,
                FieldWidthEntry {
                    fallback: field.fallback_width,
                    per_page: field.per_page.clone(),
                },
            )
        })
        .collect();
    let ctx = RenderCtx {
        page_number: page.number.unwrap_or(page_index as u64 + 1),
        page_index,
        total_pages,
        shape,
        field_widths: (!field_widths.is_empty()).then_some(&field_widths),
    };
    let content_width = page.size.w - page.margins.left - page.margins.right;
    let height = variant
        .height
        .unwrap_or_else(|| variant.measured.iter().map(measured_block_height).sum());
    let visual_top = variant.visual_top.unwrap_or(0.0);
    let visual_bottom = variant.visual_bottom.unwrap_or(height);
    let flow_height = variant.flow_height.unwrap_or(height);
    let (origin_y, flow_top) = match region.kind {
        HfKind::Header => {
            let distance = hf.header_distance.or(page.margins.header).unwrap_or(48.0);
            (distance, distance)
        }
        HfKind::Footer => {
            let distance = hf.footer_distance.or(page.margins.footer).unwrap_or(48.0);
            let actual = (visual_bottom - visual_top).max(24.0);
            (
                page.size.h - distance - actual - visual_top,
                page.size.h - distance - flow_height,
            )
        }
    };
    let geom = PageFloatGeom {
        page_width: page.size.w,
        page_height: page.size.h,
        margin_left: page.margins.left,
        margin_top: page.margins.top,
        content_width,
        content_height: page.size.h - page.margins.top - page.margins.bottom,
    };
    let mut behind = Vec::new();
    let mut flow = Vec::new();
    let mut front = Vec::new();
    let mut cursor = 0.0;
    for measured in &variant.measured {
        match (&measured.block, &measured.measure) {
            (BlockIn::Paragraph(block), MeasureIn::Paragraph(measure)) => {
                let before = block
                    .attrs
                    .as_ref()
                    .and_then(|attrs| attrs.spacing)
                    .and_then(|spacing| spacing.before)
                    .unwrap_or(0.0);
                let y = origin_y + cursor + before;
                emit_paragraph_floating_images(&mut behind, block, y, &geom, true);
                let fragment = ParagraphFragmentIn {
                    block_id: block.id.clone(),
                    x: page.margins.left,
                    y,
                    width: content_width,
                    height: measure.total_height,
                    from_line: 0,
                    to_line: measure.lines.len(),
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_paragraph_fragment(
                    &mut flow, &fragment, block, measure, &ctx, fragment.x, fragment.y, None, None,
                    true, true,
                );
                emit_paragraph_floating_images(&mut front, block, y, &geom, false);
                cursor += measure.total_height;
            }
            (BlockIn::Table(block), MeasureIn::Table(measure)) => {
                let (x, y, advances) = if let Some(floating) = &block.floating {
                    let mut top = floating.tblp_y.unwrap_or(0.0);
                    if floating.vert_anchor.as_deref() == Some("page") {
                        top -= flow_top;
                    } else if floating.vert_anchor.as_deref() == Some("margin") {
                        top += page.margins.top - flow_top;
                    }
                    let mut left = floating.tblp_x.unwrap_or(0.0);
                    if floating.horz_anchor.as_deref() == Some("page") {
                        left -= page.margins.left;
                    }
                    (page.margins.left + left, origin_y + top, false)
                } else {
                    (page.margins.left, origin_y + cursor, true)
                };
                let fragment = TableFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y,
                    width: measure.total_width,
                    height: measure.total_height,
                    row_start: 0,
                    row_end: block.rows.len(),
                    clip_top: None,
                    clip_bottom: None,
                    header_row_count: None,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_table_fragment(&mut flow, &fragment, block, measure, &ctx);
                if advances {
                    cursor += measure.total_height;
                }
            }
            (BlockIn::Image(block), MeasureIn::Image(measure)) => {
                let x = page.margins.left;
                let y = origin_y + cursor;
                let mut attrs = BlockRef::of(&block.id).attrs();
                attrs.doc_start = block.pm_start;
                attrs.doc_end = block.pm_end;
                attrs.sdt = sdt_attrs_from_groups(&block.sdt_groups);
                attrs.sdt_path = sdt_path_from_groups(&block.sdt_groups);
                stamp_image_block_attrs(&mut attrs, block, x, y);
                let rotation = block
                    .rotation_deg
                    .unwrap_or_else(|| rotation_degrees(block.transform.as_deref()));
                flow.push(Primitive::Image(ImagePrimitive {
                    rel_id: block.src.clone(),
                    x: px(x),
                    y: px(y),
                    w: px(measure.width),
                    h: px(measure.height),
                    rotation_deg: (rotation != 0.0).then(|| px(rotation)),
                    opacity: block.opacity.map(px),
                    filter: None,
                    decorative: block.decorative.unwrap_or(false),
                    crop: crop_of_block(block),
                    alt_text: capped_alt_text(block.alt.as_deref()),
                    attrs,
                }));
                cursor += measure.height;
            }
            (BlockIn::TextBox(block), MeasureIn::TextBox(measure)) => {
                let (x, y) = resolve_hf_box_position(
                    block.position.as_ref(),
                    block.css_float.as_deref(),
                    measure.width,
                    measure.height,
                    origin_y + cursor - page.margins.top,
                    &geom,
                );
                let fragment = TextBoxFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y,
                    width: measure.width,
                    height: measure.height,
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                    is_floating: Some(block.display_mode.as_deref() == Some("float")),
                    z_index: None,
                };
                emit_text_box_fragment(&mut flow, &fragment, block, measure, &ctx);
                if block.display_mode.as_deref() != Some("float")
                    && !is_floating_wrap_type(block.wrap_type.as_deref())
                {
                    cursor += measure.height;
                }
            }
            _ => {}
        }
    }
    behind.extend(flow);
    behind.extend(front);
    region.primitives = behind;
}

/// build a display list from the parsed input; the pure core the JSON boundary
/// wraps
pub fn build_display_list(input: &BuildInput, fonts: &ooxml_text::FontStore) -> DisplayList {
    build_display_list_selected(input, fonts, None)
}

fn build_display_list_selected(
    input: &BuildInput,
    fonts: &ooxml_text::FontStore,
    selected_pages: Option<&HashSet<usize>>,
) -> DisplayList {
    // Without font chains every text run stays a TextRunPrimitive.
    let shape_fonts = ShapeFonts::build(input, fonts);
    let render_options =
        serde_json::from_value::<RenderOptionsIn>(input.options.clone()).unwrap_or_default();

    // Index measured blocks by canonical block key.
    let mut by_id: HashMap<String, &MeasuredBlockIn> = HashMap::new();
    for mb in &input.measured {
        let key = match &mb.block {
            BlockIn::Paragraph(p) => block_key(&p.id),
            BlockIn::Table(t) => block_key(&t.id),
            BlockIn::Image(i) => block_key(&i.id),
            BlockIn::TextBox(t) => block_key(&t.id),
            BlockIn::Shape(s) => block_key(&s.id),
            BlockIn::Chart(c) => block_key(&c.id),
            BlockIn::Unsupported => continue,
        };
        by_id.entry(key).or_insert(mb);
    }

    let total_pages = input.layout.pages.len() as u64;
    let mut pages =
        Vec::with_capacity(selected_pages.map_or(input.layout.pages.len(), HashSet::len));

    for (page_index, page) in input.layout.pages.iter().enumerate() {
        if selected_pages.is_some_and(|selected| !selected.contains(&page_index)) {
            continue;
        }
        let ctx = RenderCtx {
            page_number: page.number.unwrap_or(page_index as u64 + 1),
            page_index,
            total_pages,
            shape: shape_fonts.as_ref(),
            field_widths: None,
        };
        let mut prims: Vec<Primitive> = Vec::new();
        let mut page_borders: Vec<PageBorderPrimitive> = Vec::new();

        if let Some(watermark) = input
            .headers_footers
            .as_ref()
            .and_then(|hf| hf.watermark.as_ref())
        {
            emit_watermark(&mut prims, watermark, page);
        }
        if let Some(border) = page_border_primitive(
            &render_options,
            page,
            page.number.unwrap_or(page_index as u64 + 1),
        ) {
            page_borders.push(border);
        }

        // Page coordinate frame for anchored-float resolution.
        let float_geom = PageFloatGeom {
            page_width: page.size.w,
            page_height: page.size.h,
            margin_left: page.margins.left,
            margin_top: page.margins.top,
            content_width: page.size.w - page.margins.left - page.margins.right,
            content_height: page.size.h - page.margins.top - page.margins.bottom,
        };

        // Behind-document floating images paint before body content.
        for frag in &page.fragments {
            if let FragmentIn::Paragraph(pf) = frag
                && let Some(mb) = by_id.get(&block_key(&pf.block_id))
                && let BlockIn::Paragraph(block) = &mb.block
            {
                emit_paragraph_floating_images(&mut prims, block, pf.y, &float_geom, true);
            }
        }

        // Paragraph border grouping uses neighboring fragment borders.
        let para_borders_of = |frag: &FragmentIn| -> Option<ParaBordersIn> {
            if let FragmentIn::Paragraph(p) = frag
                && let Some(mb) = by_id.get(&block_key(&p.block_id))
                && let BlockIn::Paragraph(b) = &mb.block
            {
                return b.attrs.as_ref().and_then(|a| a.borders.clone());
            }
            None
        };

        let mut prev_para_borders: Option<ParaBordersIn> = None;
        for (i, frag) in page.fragments.iter().enumerate() {
            match frag {
                FragmentIn::Paragraph(pf) => {
                    let Some(mb) = by_id.get(&block_key(&pf.block_id)) else {
                        prev_para_borders = None;
                        continue;
                    };
                    let (BlockIn::Paragraph(block), MeasureIn::Paragraph(measure)) =
                        (&mb.block, &mb.measure)
                    else {
                        prev_para_borders = None;
                        continue;
                    };
                    let next_borders = page.fragments.get(i + 1).and_then(para_borders_of);
                    emit_paragraph_fragment(
                        &mut prims,
                        pf,
                        block,
                        measure,
                        &ctx,
                        pf.x,
                        pf.y,
                        prev_para_borders.as_ref(),
                        next_borders.as_ref(),
                        true,
                        // body fragments surface as mirror paragraph wrappers
                        true,
                    );
                    prev_para_borders = block.attrs.as_ref().and_then(|a| a.borders.clone());
                }
                FragmentIn::Table(tf) => {
                    prev_para_borders = None;
                    let Some(mb) = by_id.get(&block_key(&tf.block_id)) else {
                        continue;
                    };
                    let (BlockIn::Table(block), MeasureIn::Table(measure)) =
                        (&mb.block, &mb.measure)
                    else {
                        continue;
                    };
                    emit_table_fragment(&mut prims, tf, block, measure, &ctx);
                }
                FragmentIn::Image(imf) => {
                    prev_para_borders = None;
                    let block = by_id.get(&block_key(&imf.block_id)).and_then(|mb| {
                        if let BlockIn::Image(b) = &mb.block {
                            Some(b)
                        } else {
                            None
                        }
                    });
                    let mut attrs = BlockRef::of(&imf.block_id).attrs();
                    attrs.doc_start = imf.pm_start.or(block.and_then(|b| b.pm_start));
                    attrs.doc_end = imf.pm_end.or(block.and_then(|b| b.pm_end));
                    attrs.sdt = block.and_then(|b| sdt_attrs_from_groups(&b.sdt_groups));
                    if let Some(block) = block {
                        attrs.sdt_path = sdt_path_from_groups(&block.sdt_groups);
                        stamp_image_block_attrs(&mut attrs, block, imf.x, imf.y);
                    }
                    let rot = block.and_then(|b| b.rotation_deg).unwrap_or_else(|| {
                        rotation_degrees(block.and_then(|b| b.transform.as_deref()))
                    });
                    prims.push(Primitive::Image(ImagePrimitive {
                        rel_id: block.map(|b| b.src.clone()).unwrap_or_default(),
                        x: px(imf.x),
                        y: px(imf.y),
                        w: px(imf.width),
                        h: px(imf.height),
                        rotation_deg: if rot != 0.0 { Some(px(rot)) } else { None },
                        opacity: block.and_then(|b| b.opacity).map(px),
                        filter: None,
                        decorative: block.and_then(|b| b.decorative).unwrap_or(false),
                        crop: block.and_then(|block| crop_of_block(block)),
                        alt_text: capped_alt_text(block.and_then(|b| b.alt.as_deref())),
                        attrs,
                    }));
                }
                FragmentIn::TextBox(tf) => {
                    prev_para_borders = None;
                    let Some(mb) = by_id.get(&block_key(&tf.block_id)) else {
                        continue;
                    };
                    let (BlockIn::TextBox(block), MeasureIn::TextBox(measure)) =
                        (&mb.block, &mb.measure)
                    else {
                        continue;
                    };
                    emit_text_box_fragment(&mut prims, tf, block, measure, &ctx);
                }
                FragmentIn::Shape(sf) => {
                    prev_para_borders = None;
                    let Some(mb) = by_id.get(&block_key(&sf.block_id)) else {
                        continue;
                    };
                    let BlockIn::Shape(block) = &mb.block else {
                        continue;
                    };
                    emit_shape_fragment(&mut prims, sf, block, &ctx);
                }
                FragmentIn::Chart(cf) => {
                    prev_para_borders = None;
                    let Some(mb) = by_id.get(&block_key(&cf.block_id)) else {
                        continue;
                    };
                    let BlockIn::Chart(block) = &mb.block else {
                        continue;
                    };
                    emit_chart_fragment(&mut prims, cf, block);
                }
                FragmentIn::Unsupported => {
                    prev_para_borders = None;
                }
            }
        }

        // Front floating images paint after body content.
        for frag in &page.fragments {
            if let FragmentIn::Paragraph(pf) = frag
                && let Some(mb) = by_id.get(&block_key(&pf.block_id))
                && let BlockIn::Paragraph(block) = &mb.block
            {
                emit_paragraph_floating_images(&mut prims, block, pf.y, &float_geom, false);
            }
        }

        emit_column_separators(&mut prims, page);

        // header/footer bands: resolve the page's variant and compose its
        // region with the same emitters (see hf_bands.rs for the geometry)
        let (mut header, mut footer) = match &input.headers_footers {
            Some(hf) => crate::hf_bands::compose_page_regions(
                hf,
                page,
                page_index,
                total_pages,
                shape_fonts.as_ref(),
            ),
            None => (None, None),
        };
        if let Some(hf) = &input.headers_footers_content {
            if let Some(region) = &mut header {
                recompose_hf_region(
                    region,
                    hf,
                    page,
                    page_index,
                    total_pages,
                    shape_fonts.as_ref(),
                );
            }
            if let Some(region) = &mut footer {
                recompose_hf_region(
                    region,
                    hf,
                    page,
                    page_index,
                    total_pages,
                    shape_fonts.as_ref(),
                );
            }
        }
        let note_areas = emit_note_regions(page, &ctx);
        let (content_bounds, column_bounds) = page_content_geometry(page);

        pages.push(DisplayPage {
            page_index: page_index as u64,
            width: px(page.size.w),
            height: px(page.size.h),
            content_bounds: Some(content_bounds),
            column_bounds,
            section_id: page.section_id.clone(),
            section_index: page.section_index,
            section_page_index: page.section_page_index,
            section_page_number: page.section_page_number,
            page_label: page.page_label.clone(),
            primitives: prims,
            background: page.background.clone(),
            page_borders,
            header,
            footer,
            note_areas,
        });
    }

    let mut display_list = DisplayList {
        contract_version: input.contract_version,
        pages,
    };
    apply_review_metadata(
        &mut display_list,
        &input.resolved_comment_ids,
        &input.comment_authors,
        &input.comment_threads,
    );
    display_list
}

fn reviewer_palette_color(index: u64) -> &'static str {
    const COLORS: [&str; 8] = [
        "#c67c00", "#1565c0", "#6a1b9a", "#00838f", "#ad1457", "#2e7d32", "#5d4037", "#455a64",
    ];
    COLORS[index as usize % COLORS.len()]
}

fn comment_author_for_ids<'a>(
    ids: &[String],
    authors: &'a [CommentAuthorIn],
) -> Option<&'a CommentAuthorIn> {
    authors
        .iter()
        .find(|author| {
            author
                .id
                .as_ref()
                .is_some_and(|author_id| ids.iter().any(|id| id == author_id))
        })
        // The frozen input has no comment-id → author-id join. A single
        // reviewer is unambiguous; with multiple unmatched reviewers, omit
        // attribution rather than assigning the wrong person's identity.
        .or_else(|| (authors.len() == 1).then(|| &authors[0]))
}

/// comment-id keyed thread lookup: the first thread whose id matches any of
/// the primitive's comment ids (Word ranges rarely overlap; when they do the
/// first/outermost id wins, matching the wash color choice)
fn comment_thread_for_ids<'a>(
    ids: &[String],
    threads: &'a [CommentThreadIn],
) -> Option<&'a CommentThreadIn> {
    ids.iter()
        .filter_map(|id| id.parse::<i64>().ok())
        .find_map(|comment_id| threads.iter().find(|thread| thread.id == Some(comment_id)))
}

/// comment body/reply text cap — announcement excerpts, not full transcripts
pub const MAX_COMMENT_TEXT_CHARS: usize = 512;
/// reply-summary cap; `reply_count` still reports the real total
pub const MAX_COMMENT_REPLIES: usize = 16;

/// bounded copy of a file-derived string: empty drops, oversized truncates on
/// a char boundary (same policy as [`capped_alt_text`])
fn capped_string(value: Option<&str>, max_chars: usize) -> Option<String> {
    let value = value?;
    if value.is_empty() {
        return None;
    }
    Some(value.chars().take(max_chars).collect())
}

/// merge one thread's a11y metadata into the primitive's CommentMetadata:
/// author name/date/body plus bounded reply summaries. Everything here is
/// file-derived and announce-only — the mirror assigns it via
/// setAttribute/textContent, never markup.
fn apply_comment_thread_metadata(metadata: &mut CommentMetadata, thread: &CommentThreadIn) {
    metadata.author_name = capped_string(thread.author_name.as_deref(), MAX_COMMENT_TEXT_CHARS);
    metadata.date = capped_string(thread.date.as_deref(), MAX_COMMENT_TEXT_CHARS);
    metadata.text = capped_string(thread.text.as_deref(), MAX_COMMENT_TEXT_CHARS);
    metadata.reply_count = Some(thread.replies.len() as u64);
    metadata.replies = thread
        .replies
        .iter()
        .take(MAX_COMMENT_REPLIES)
        .map(|reply| CommentReplyMetadata {
            author_name: capped_string(reply.author_name.as_deref(), MAX_COMMENT_TEXT_CHARS),
            date: capped_string(reply.date.as_deref(), MAX_COMMENT_TEXT_CHARS),
            text: capped_string(reply.text.as_deref(), MAX_COMMENT_TEXT_CHARS),
        })
        .collect();
    // the thread input carries the real comment-id → author join the frozen
    // author palette lacked; fill authorId only when the palette heuristic
    // found none, so existing attribution stays byte-stable
    if metadata.author_id.is_none() {
        metadata.author_id = capped_string(thread.author_id.as_deref(), MAX_COMMENT_TEXT_CHARS);
    }
}

fn review_color(author: &CommentAuthorIn) -> String {
    author
        .color
        .clone()
        .unwrap_or_else(|| reviewer_palette_color(author.palette_index.unwrap_or(0)).to_string())
}

fn comment_wash(color: &str) -> String {
    let hex = color.strip_prefix('#').unwrap_or(color);
    if hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        let red = u8::from_str_radix(&hex[0..2], 16).unwrap_or(255);
        let green = u8::from_str_radix(&hex[2..4], 16).unwrap_or(212);
        let blue = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
        format!("rgba({red}, {green}, {blue}, 0.15)")
    } else {
        color.to_string()
    }
}

fn apply_review_primitive_metadata(
    primitives: &mut Vec<Primitive>,
    resolved_ids: &HashSet<i64>,
    authors: &[CommentAuthorIn],
    threads: &[CommentThreadIn],
) {
    primitives.retain_mut(|primitive| {
        let is_comment_wash = matches!(
            primitive,
            Primitive::Decoration(DecorationPrimitive {
                deco: DecoKind::CommentRange,
                ..
            })
        );
        let comment_ids = doc_attrs_mut(primitive).and_then(|attrs| attrs.comment_ids.clone());
        if let Some(comment_ids) = comment_ids {
            let all_resolved = !comment_ids.is_empty()
                && comment_ids.iter().all(|id| {
                    id.parse::<i64>()
                        .is_ok_and(|comment_id| resolved_ids.contains(&comment_id))
                });
            let author = comment_author_for_ids(&comment_ids, authors);
            let color = author.map(review_color);
            let thread = comment_thread_for_ids(&comment_ids, threads);
            if let Some(attrs) = doc_attrs_mut(primitive) {
                let mut metadata = CommentMetadata {
                    status: Some(if all_resolved { "resolved" } else { "active" }.to_string()),
                    author_id: author.and_then(|entry| entry.id.clone()),
                    palette_index: author.and_then(|entry| entry.palette_index),
                    color: color.clone(),
                    selected: None,
                    ..CommentMetadata::default()
                };
                if let Some(thread) = thread {
                    apply_comment_thread_metadata(&mut metadata, thread);
                }
                attrs.comment = Some(metadata);
            }
            if is_comment_wash {
                if all_resolved {
                    return false;
                }
                if let Some(color) = color
                    && let Primitive::Decoration(decoration) = &mut *primitive
                {
                    decoration.color = comment_wash(&color);
                }
            }
        }

        // Revisions do carry an author name, so author-colored underline and
        // strike recipes can be applied without the missing comment join.
        if let Primitive::Decoration(decoration) = primitive
            && matches!(decoration.deco, DecoKind::Underline | DecoKind::Strike)
            && let Some(revision) = &decoration.attrs.revision
            && let Some(author) = authors.iter().find(|entry| {
                entry.name.as_deref() == Some(revision.author.as_str())
                    || entry.id.as_deref() == Some(revision.author.as_str())
            })
        {
            decoration.color = review_color(author);
        }
        true
    });
}

fn apply_review_metadata(
    display_list: &mut DisplayList,
    resolved_comment_ids: &[i64],
    authors: &[CommentAuthorIn],
    threads: &[CommentThreadIn],
) {
    let resolved_ids: HashSet<i64> = resolved_comment_ids.iter().copied().collect();
    for page in &mut display_list.pages {
        apply_review_primitive_metadata(&mut page.primitives, &resolved_ids, authors, threads);
        if let Some(header) = &mut page.header {
            apply_review_primitive_metadata(
                &mut header.primitives,
                &resolved_ids,
                authors,
                threads,
            );
        }
        if let Some(footer) = &mut page.footer {
            apply_review_primitive_metadata(
                &mut footer.primitives,
                &resolved_ids,
                authors,
                threads,
            );
        }
        for area in &mut page.note_areas {
            apply_review_primitive_metadata(
                &mut area.separator_primitives,
                &resolved_ids,
                authors,
                threads,
            );
            apply_review_primitive_metadata(&mut area.primitives, &resolved_ids, authors, threads);
        }
    }
}

/// Paints one paragraph fragment: shading rect, grouped borders, then each
/// measured line's runs with indent padding, per-line float offsets and the
/// alignment shift.
///
/// `origin_*` are the fragment's page coordinates; cell content passes its own
/// origin and suppresses border grouping. `stamp_line_range` records the
/// fragment's `[from_line, to_line)` on every primitive it emits — table-cell
/// callers pass `false`, since a cell paragraph has no fragment of its own for
/// the range to describe.
#[allow(clippy::too_many_arguments)]
pub(crate) fn emit_paragraph_fragment(
    prims: &mut Vec<Primitive>,
    frag: &ParagraphFragmentIn,
    block: &ParagraphBlockIn,
    measure: &ParagraphExtentIn,
    ctx: &RenderCtx<'_>,
    origin_x: f64,
    origin_y: f64,
    prev_borders: Option<&ParaBordersIn>,
    next_borders: Option<&ParaBordersIn>,
    emit_block_chrome: bool,
    stamp_line_range: bool,
) {
    // primitives emitted for this fragment get the paragraph's paraId stamped
    // afterward (mirror `data-para-id`); remember where they start
    let stamp_from = prims.len();
    let attrs = block.attrs.as_ref();
    let is_rtl = attrs.and_then(|a| a.bidi).unwrap_or(false) || paragraph_base_is_rtl(block);
    let indent = attrs.and_then(|a| a.indent).unwrap_or_default();

    // rtl paragraphs mirror the two indent sides
    let ind_l = indent.left.unwrap_or(0.0).max(0.0);
    let ind_r = indent.right.unwrap_or(0.0).max(0.0);
    let (indent_left, indent_right) = if is_rtl {
        (ind_r, ind_l)
    } else {
        (ind_l, ind_r)
    };
    let hanging = indent.hanging.unwrap_or(0.0).max(0.0);
    let first_line = indent.first_line.unwrap_or(0.0).max(0.0);

    let block_ref = BlockRef::of(&frag.block_id);
    let pmark_revision = attrs.and_then(|a| {
        a.p_pr_ins
            .as_ref()
            .map(|info| {
                structural_revision(
                    info,
                    StructuralRevisionScope::Pmark,
                    StructuralRevisionKind::Ins,
                    None,
                    None,
                )
            })
            .or_else(|| {
                a.p_pr_del.as_ref().map(|info| {
                    structural_revision(
                        info,
                        StructuralRevisionScope::Pmark,
                        StructuralRevisionKind::Del,
                        None,
                        None,
                    )
                })
            })
    });

    if emit_block_chrome {
        if let Some(rev) = pmark_revision.clone() {
            let mut bar_attrs = block_ref.attrs();
            bar_attrs.structural_revision = Some(rev.clone());
            prims.push(Primitive::Rect(RectPrimitive {
                x: px(origin_x + STRUCTURAL_CHANGE_BAR_OFFSET_X),
                y: px(origin_y),
                w: px(STRUCTURAL_CHANGE_BAR_WIDTH),
                h: px(frag.height),
                fill: structural_color(rev.kind).to_string(),
                attrs: bar_attrs,
            }));
        }
        // paragraph shading paints as a fragment-sized rect behind everything
        if let Some(shading) = attrs.and_then(|a| a.shading.as_ref()) {
            let mut shading_attrs = block_ref.attrs();
            shading_attrs.doc_start = frag.pm_start;
            shading_attrs.doc_end = frag.pm_end;
            prims.push(Primitive::Rect(RectPrimitive {
                x: px(origin_x),
                y: px(origin_y),
                w: px(frag.width),
                h: px(frag.height),
                fill: shading.clone(),
                attrs: shading_attrs,
            }));
        }
        if let Some(borders) = attrs.and_then(|a| a.borders.as_ref()) {
            emit_paragraph_borders(
                prims,
                borders,
                prev_borders,
                next_borders,
                origin_x,
                origin_y,
                frag.width,
                frag.height,
                indent_left,
                indent_right,
            );
        }
    }

    let total_lines = measure.lines.len();
    let carried_from_prev = frag.carried_from_prev == Some(true);
    let alignment = attrs.and_then(|a| a.alignment.as_deref());
    // A visible list marker occupies the first-line hanging indent.
    let has_list_marker = attrs
        .and_then(|a| a.list_marker.as_deref())
        .is_some_and(|m| !m.is_empty())
        && attrs.and_then(|a| a.list_marker_hidden) != Some(true);
    // a trailing <w:br> makes Word justify the paragraph's final line too
    let para_ends_with_line_break = matches!(block.runs.last(), Some(RunIn::LineBreak(_)));

    let mut line_top = origin_y;
    let mut last_line: Option<LinePaintMetrics> = None;
    for idx in frag.from_line..frag.to_line.min(total_lines) {
        let line = &measure.lines[idx];
        // lead skip reserves vertical space above the line (float push-down)
        line_top += line.float_skip_before.unwrap_or(0.0);

        let is_first_line = idx == 0 && !carried_from_prev;
        let is_last_line = idx + 1 == total_lines;
        let line_stamp_from = prims.len();

        last_line = emit_line(
            prims,
            block,
            line,
            LineGeom {
                frag_x: origin_x,
                frag_width: frag.width,
                line_top,
                indent_left,
                indent_right,
                hanging,
                first_line,
                is_first_line,
                is_last_line,
                is_rtl,
                alignment,
                frag_pm_start: frag.pm_start,
                has_list_marker,
                para_ends_with_line_break,
            },
            &block_ref,
            ctx,
        );
        for primitive in &mut prims[line_stamp_from..] {
            if let Some(attrs) = doc_attrs_mut(primitive) {
                attrs.line_index = Some(idx as u64);
            }
        }

        line_top += line.line_height;
    }

    // The tracked-change pilcrow marks the paragraph's end, so only the final
    // fragment gets it; earlier fragments get the margin bar alone.
    if let (Some(rev), Some(line)) = (pmark_revision, last_line)
        && frag.carried_to_next != Some(true)
    {
        let mut glyph_attrs = block_ref.attrs();
        glyph_attrs.structural_revision = Some(rev.clone());
        let glyph_x = line.end_x + PARAGRAPH_MARK_GLYPH_GAP;
        prims.push(Primitive::Text(TextRunPrimitive {
            text: "¶".to_string(),
            x: px(glyph_x),
            baseline_y: px(line.baseline),
            width: px(PARAGRAPH_MARK_GLYPH_WIDTH),
            paint_clip: None,
            font: css_font(&RunFormattingIn::default()),
            color: structural_color(rev.kind).to_string(),
            letter_spacing: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: glyph_attrs.clone(),
        }));
        if rev.kind == StructuralRevisionKind::Del {
            prims.push(Primitive::Decoration(DecorationPrimitive {
                deco: DecoKind::Strike,
                x: px(glyph_x),
                y: px(line.baseline - (font_px_of(&RunFormattingIn::default()) * 0.3).round()),
                w: px(PARAGRAPH_MARK_GLYPH_WIDTH),
                h: px(1.0),
                color: structural_color(rev.kind).to_string(),
                dashed: false,
                dotted: false,
                attrs: glyph_attrs,
            }));
        }
    }

    // stamp the paragraph's stable paraId AND (for mirror-surfaced fragments) its
    // measured line window on every primitive it emitted, so the a11y mirror can
    // group the wrapper and expose `data-para-id` / `data-from-line` /
    // `data-to-line` (lines carry no DocAttrs and are left untouched)
    let para_id = block.para_id.as_ref();
    let line_range = if stamp_line_range {
        Some((frag.from_line as u64, frag.to_line as u64))
    } else {
        None
    };
    let sdt = sdt_attrs_from_groups(&block.sdt_groups);
    let sdt_path = sdt_path_from_groups(&block.sdt_groups);
    if para_id.is_some()
        || line_range.is_some()
        || sdt.is_some()
        || frag.pm_start.is_some()
        || frag.pm_end.is_some()
    {
        for p in &mut prims[stamp_from..] {
            if let Some(a) = doc_attrs_mut(p) {
                a.fragment_doc_start = frag.pm_start;
                a.fragment_doc_end = frag.pm_end;
                if let Some(pid) = para_id {
                    a.para_id = Some(pid.clone());
                }
                if let Some((from, to)) = line_range {
                    a.from_line = Some(from);
                    a.to_line = Some(to);
                }
                if let Some(sdt) = &sdt {
                    a.sdt = Some(sdt.clone());
                    a.sdt_path = sdt_path.clone();
                }
            }
        }
    }
}

/// Everything one line needs beyond its own measure: the fragment box it sits
/// in, the paragraph's indents, its position within the paragraph, and the
/// direction and alignment that place its content.
struct LineGeom<'a> {
    frag_x: f64,
    frag_width: f64,
    line_top: f64,
    indent_left: f64,
    indent_right: f64,
    hanging: f64,
    first_line: f64,
    is_first_line: bool,
    is_last_line: bool,
    is_rtl: bool,
    alignment: Option<&'a str>,
    frag_pm_start: Option<i64>,
    /// Whether the first line reserves the hanging indent for a list marker.
    has_list_marker: bool,
    /// A trailing explicit line break allows the closing line to justify.
    para_ends_with_line_break: bool,
}

#[derive(Clone, Copy)]
struct LinePaintMetrics {
    end_x: f64,
    baseline: f64,
}

/// Materialize the authoritative Rust-measure seam into paint items. Every
/// tuple uses the measure engine's UTF-16 cluster boundary and exact advance;
/// no width is inferred from the number of Unicode scalars. `bidiSlices` wins
/// because it carries explicit visual order, followed by cluster metadata and
/// finally exact run slices for older authoritative payloads.
fn authoritative_line_items<'a>(
    block: &'a ParagraphBlockIn,
    line: &LineIn,
    ctx: &RenderCtx<'_>,
    default_level: u8,
) -> Option<Vec<LinePaintItem<'a>>> {
    let mut slices: Vec<(usize, usize, usize, f64, u8, u64, Option<u64>)> = Vec::new();
    if !line.bidi_slices.is_empty() {
        for (index, slice) in line.bidi_slices.iter().enumerate() {
            slices.push((
                slice.run_index?,
                slice.start_char?,
                slice.end_char?,
                slice.advance?,
                slice.bidi_level.unwrap_or(default_level),
                slice.visual_order.unwrap_or(index as u64),
                slice.logical_order,
            ));
        }
        slices.sort_by_key(|slice| slice.5);
    } else if !line.cluster_advances.is_empty() {
        for (index, cluster) in line.cluster_advances.iter().enumerate() {
            slices.push((
                cluster.run_index?,
                cluster.start_char?,
                cluster.end_char?,
                cluster.advance?,
                cluster.bidi_level.unwrap_or(default_level),
                // xOffset is the authoritative visual coordinate. Convert it
                // to a stable integer sort key without changing the value.
                cluster
                    .x_offset
                    .map(|x| (x.max(0.0) * 1_000.0).round() as u64)
                    .unwrap_or(index as u64),
                cluster.logical_order,
            ));
        }
        slices.sort_by_key(|slice| slice.5);
    } else if !line.run_advances.is_empty() {
        for (index, run) in line.run_advances.iter().enumerate() {
            let run_index = run.run_index?;
            let level = match block.runs.get(run_index) {
                Some(RunIn::Text(text)) => text.fmt.bidi_level.unwrap_or(default_level),
                Some(RunIn::Field(field)) => field.fmt.bidi_level.unwrap_or(default_level),
                Some(RunIn::Tab(tab)) => tab.fmt.bidi_level.unwrap_or(default_level),
                _ => default_level,
            };
            slices.push((
                run_index,
                run.start_char?,
                run.end_char?,
                run.advance?,
                level,
                index as u64,
                run.logical_order,
            ));
        }
    } else {
        return None;
    }

    let item_count = slices.len();
    let mut out = Vec::with_capacity(item_count);
    for (run_index, start, end, measured_width, level, _, logical_order) in slices {
        let run = block.runs.get(run_index)?;
        match run {
            RunIn::Text(text) => {
                let visible = slice_utf16(&text.text, start, end);
                out.push(LinePaintItem::Text(LineTextItem {
                    text: visible,
                    fmt: &text.fmt,
                    pm_start: text.pm_start.map(|pos| pos + start as i64),
                    pm_end: text.pm_start.map(|pos| pos + end as i64).or(text.pm_end),
                    width: measured_width,
                    level,
                    source_start: start,
                    source_end: end,
                    logical_order: logical_order.or(text.fmt.logical_order),
                    exact_advance: true,
                    field: None,
                }));
            }
            RunIn::Field(field) => {
                let text = field_text(field, ctx);
                let width = ctx
                    .field_width(field.pm_start)
                    .map(|(_, resolved)| resolved)
                    .unwrap_or(measured_width);
                out.push(LinePaintItem::Text(LineTextItem {
                    text,
                    fmt: &field.fmt,
                    pm_start: field.pm_start,
                    pm_end: field.pm_end,
                    width,
                    level,
                    source_start: start,
                    source_end: end,
                    logical_order: logical_order.or(field.fmt.logical_order),
                    exact_advance: true,
                    field: Some(field),
                }));
            }
            RunIn::Tab(tab) => out.push(LinePaintItem::Tab {
                run: tab,
                width: measured_width,
                level,
                logical_order,
            }),
            RunIn::Image(image) if !is_floating_image_run(image) => {
                out.push(LinePaintItem::Image {
                    run: image,
                    pm_start: image.pm_start,
                    pm_end: image.pm_end,
                    single_image_line: item_count == 1,
                    level,
                    logical_order,
                });
            }
            RunIn::LineBreak(line_break) => out.push(LinePaintItem::LineBreak {
                pm_start: line_break.pm_start,
                level,
            }),
            RunIn::Image(_) | RunIn::Unsupported => {}
        }
    }
    Some(out)
}

/// paint one typeset line: pen starts at the indent-derived x, runs advance by
/// their resolved widths (tabs/images exact, text distributed over the measured
/// line width), decorations ride with their run in paint order (highlight and
/// comment tint behind the glyphs, underline/strike after).
fn emit_line(
    prims: &mut Vec<Primitive>,
    block: &ParagraphBlockIn,
    line: &LineIn,
    geom: LineGeom,
    block_ref: &BlockRef,
    ctx: &RenderCtx<'_>,
) -> Option<LinePaintMetrics> {
    let segments = resolve_line_segments(&block.runs, line);
    let attrs = block.attrs.as_ref();
    let default_bidi_level = base_bidi_level(geom.is_rtl);
    let authoritative_items = authoritative_line_items(block, line, ctx, default_bidi_level);
    let authoritative_active = authoritative_items.is_some();
    let authoritative_width = authoritative_items
        .as_ref()
        .map(|items| items.iter().map(LinePaintItem::width).sum::<f64>());

    // The first line carries the hanging or first-line shift.
    let has_hanging = geom.hanging > 0.0;
    let has_first_line = geom.first_line > 0.0;
    let mut pad_left = geom.indent_left;
    let mut text_indent = 0.0;
    if geom.is_first_line {
        if geom.indent_left > 0.0 && has_hanging {
            text_indent = if geom.has_list_marker {
                // The marker consumes the hanging width before body text.
                (geom.hanging - geom.indent_left).max(0.0)
            } else {
                -geom.hanging
            };
        } else if has_first_line {
            text_indent = geom.first_line;
        } else if geom.has_list_marker && has_hanging {
            // no left indent: the marker fills [0, hanging] and body text
            // follows at the hang (matches the body-line branch below)
            text_indent = geom.hanging;
        }
    } else if geom.indent_left <= 0.0 && has_hanging {
        pad_left = geom.hanging;
    }

    let left_offset = line.left_offset.unwrap_or(0.0);
    let right_offset = line.right_offset.unwrap_or(0.0);
    let usable_width =
        (geom.frag_width - pad_left - geom.indent_right - left_offset - right_offset).max(0.0);

    // distribute the measured line width over runs whose advance we don't know
    // individually: fixed-width runs (tabs, inline images) subtract first, the
    // rest splits across text-ish segments by their fallback slots.
    //
    // Per-page field widths are fixed and removed from the text pool.
    let mut fixed_width = 0.0;
    let default_font_pt = attrs
        .and_then(|attrs| attrs.default_font_size)
        .filter(|size| size.is_finite() && *size > 0.0)
        .unwrap_or(DEFAULT_FONT_PT);
    let mut pool_estimate = 0.0;
    let mut field_fixed = 0.0; // Σ per-page resolved widths of supplied fields
    let mut field_fallback = 0.0; // Σ their fallback widths baked into line.width
    if authoritative_items.is_none() {
        for seg in &segments {
            match seg.run {
                RunIn::Tab(t) => fixed_width += t.width.unwrap_or(48.0),
                RunIn::Image(imr) => {
                    if !is_floating_image_run(imr) {
                        fixed_width += image_layout_width(imr);
                    }
                }
                RunIn::Text(text) => {
                    pool_estimate += fallback_text_width(&seg.text, &text.fmt, default_font_pt)
                }
                RunIn::Field(f) => match ctx.field_width(seg.pm_start) {
                    Some((fallback, resolved)) => {
                        field_fallback += fallback;
                        field_fixed += resolved;
                    }
                    None => {
                        pool_estimate +=
                            fallback_text_width(&field_text(f, ctx), &f.fmt, default_font_pt)
                    }
                },
                _ => {}
            }
        }
    }
    let pool_width = (line.width - fixed_width - field_fallback).max(0.0);
    let width_per_estimate = if pool_estimate > 0.0 {
        pool_width / pool_estimate
    } else {
        0.0
    };
    // the line's true rendered extent on this page; equals line.width when no
    // per-page field widths were supplied
    let effective_line_width =
        authoritative_width.unwrap_or(pool_width + fixed_width + field_fixed);

    // Justified lines already fill the width, so they take no alignment shift.
    let align_shift = match geom.alignment {
        Some("center") => ((usable_width - effective_line_width) / 2.0).max(0.0),
        Some("right") => (usable_width - effective_line_width).max(0.0),
        None if geom.is_rtl => (usable_width - effective_line_width).max(0.0),
        _ => 0.0,
    };
    // suppress the stretch shift on the paragraph's un-stretched closing line
    let align_shift = if geom.alignment == Some("justify") && !geom.is_last_line {
        0.0
    } else {
        align_shift
    };

    let mut pen_x = geom.frag_x + pad_left + text_indent + left_offset + align_shift;
    let line_bottom = geom.line_top + line.line_height;
    // baseline from the measured metrics with CSS half-leading centering
    let half_leading = ((line.line_height - line.ascent - line.descent) / 2.0).max(0.0);
    let baseline = geom.line_top + half_leading + line.ascent;

    // Numbering is not part of the story text, so materialize the precomputed
    // marker as its own first-line primitive. The hanging-indent slot is its
    // authoritative horizontal extent; body text already begins after it.
    if geom.is_first_line
        && let Some(marker) = attrs
            .and_then(|attrs| attrs.list_marker.as_deref())
            .filter(|marker| !marker.is_empty())
        && attrs.and_then(|attrs| attrs.list_marker_hidden) != Some(true)
    {
        let mut marker_format = RunFormattingIn::default();
        marker_format.font_family = attrs.and_then(|attrs| {
            attrs
                .list_marker_font_family
                .clone()
                .or_else(|| attrs.default_font_family.clone())
        });
        marker_format.font_size =
            attrs.and_then(|attrs| attrs.list_marker_font_size.or(attrs.default_font_size));
        let slot_width = geom.hanging.max(font_px_of(&marker_format));
        let marker_x = if geom.is_rtl {
            pen_x + effective_line_width
        } else {
            pen_x - slot_width
        };
        let marker_revision = attrs.and_then(|attrs| attrs.list_marker_revision);
        let color = match marker_revision {
            Some(RevisionKind::Ins) => REVISION_INS_COLOR,
            Some(RevisionKind::Del) => REVISION_DEL_COLOR,
            None => "#000000",
        };
        let mut marker_attrs = block_ref.attrs();
        marker_attrs.list_marker = Some(true);
        marker_attrs.list_marker_revision = marker_revision;
        prims.push(Primitive::Text(TextRunPrimitive {
            text: marker.to_owned(),
            x: px(marker_x),
            baseline_y: px(baseline),
            width: px(slot_width),
            paint_clip: None,
            font: css_font(&marker_format),
            color: color.to_owned(),
            letter_spacing: None,
            word_spacing: None,
            rtl: geom.is_rtl.then_some(true),
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: marker_attrs,
        }));
    }

    // Distribute justification slack across U+0020 clusters.
    let justified = geom.alignment == Some("justify")
        && ooxml_text::line_is_justified(
            geom.is_last_line,
            geom.para_ends_with_line_break,
            &ooxml_text::CompatFlags::default(),
        );
    let mut word_space_px: Option<Number> = None;
    let mut word_space_extra = 0.0_f64;
    if justified && (pool_estimate > 0.0 || authoritative_items.is_some()) {
        let slack = (usable_width - effective_line_width).max(0.0);
        if slack > 0.0 {
            let mut space_count = 0usize;
            if let Some(items) = &authoritative_items {
                space_count = items
                    .iter()
                    .filter(|item| matches!(item, LinePaintItem::Text(text) if text.text == " "))
                    .count();
            } else {
                for seg in &segments {
                    let text = match seg.run {
                        RunIn::Text(_) => seg.text.clone(),
                        // a field with a supplied per-page width is fixed-width, so
                        // it's out of the text pool and never absorbs justify slack
                        RunIn::Field(f) if ctx.field_width(seg.pm_start).is_none() => {
                            field_text(f, ctx)
                        }
                        _ => continue,
                    };
                    space_count += text.chars().filter(|&ch| ch == ' ').count();
                }
            }
            if space_count > 0 {
                word_space_extra = slack / space_count as f64;
                word_space_px = Some(px(word_space_extra));
            }
        }
    }

    let mut emitted_positioned_text = false;
    let mut line_break_pos: Option<i64> = None;

    let bidi_base = base_bidi_direction(geom.is_rtl);
    let mut line_text = String::new();
    for seg in &segments {
        match seg.run {
            RunIn::Text(_) => line_text.push_str(&seg.text),
            RunIn::Field(f) => line_text.push_str(&field_text(f, ctx)),
            _ => {}
        }
    }
    let line_levels = bidi_char_levels(&line_text, bidi_base);
    let mut level_cursor = 0usize;
    let mut logical_items: Vec<LinePaintItem<'_>> = authoritative_items.unwrap_or_default();

    if logical_items.is_empty() {
        for seg in &segments {
            match seg.run {
                // hidden runs (w:vanish) stay in the doc-position flow — the DOM
                // editing view paints them dimmed rather than suppressing them, so
                // the display list keeps their primitives for hit-testing too
                RunIn::Text(t) => {
                    let w = width_per_estimate
                        * fallback_text_width(&seg.text, &t.fmt, default_font_pt)
                        + word_space_extra
                            * seg.text.chars().filter(|&ch| ch == ' ').count() as f64;
                    push_bidi_text_items(
                        &mut logical_items,
                        &seg.text,
                        &t.fmt,
                        seg.pm_start,
                        seg.pm_end,
                        w,
                        &line_levels,
                        &mut level_cursor,
                        default_bidi_level,
                        word_space_extra,
                        None,
                    );
                }
                RunIn::Field(f) => {
                    let text = field_text(f, ctx);
                    // Supplied field widths override character distribution.
                    let (w, item_word_space_extra) = match ctx.field_width(seg.pm_start) {
                        Some((_, resolved)) => (resolved, 0.0),
                        None => (
                            width_per_estimate
                                * fallback_text_width(&text, &f.fmt, default_font_pt)
                                + word_space_extra
                                    * text.chars().filter(|&ch| ch == ' ').count() as f64,
                            word_space_extra,
                        ),
                    };
                    push_bidi_text_items(
                        &mut logical_items,
                        &text,
                        &f.fmt,
                        seg.pm_start,
                        seg.pm_end,
                        w,
                        &line_levels,
                        &mut level_cursor,
                        default_bidi_level,
                        item_word_space_extra,
                        Some(f),
                    );
                }
                RunIn::Tab(t) => {
                    logical_items.push(LinePaintItem::Tab {
                        run: t,
                        width: t.width.unwrap_or(48.0),
                        level: default_bidi_level,
                        logical_order: t.fmt.logical_order,
                    });
                }
                RunIn::Image(imr) => {
                    // The page and cell float layers own these, not the line.
                    if is_floating_image_run(imr) {
                        continue;
                    }
                    logical_items.push(LinePaintItem::Image {
                        run: imr,
                        pm_start: seg.pm_start,
                        pm_end: seg.pm_end,
                        single_image_line: segments.len() == 1,
                        level: default_bidi_level,
                        logical_order: None,
                    });
                }
                RunIn::LineBreak(lb) => {
                    logical_items.push(LinePaintItem::LineBreak {
                        pm_start: lb.pm_start,
                        level: default_bidi_level,
                    });
                }
                RunIn::Unsupported => {}
            }
        }
    }

    let visual_order = if !authoritative_active {
        let item_levels: Vec<u8> = logical_items.iter().map(LinePaintItem::level).collect();
        ooxml_text::visual_order_for_levels(&item_levels)
    } else {
        (0..logical_items.len()).collect()
    };

    for item_idx in visual_order {
        let Some(item) = logical_items.get(item_idx) else {
            continue;
        };
        match item {
            LinePaintItem::Text(item) => {
                let paint_width = item.width
                    + if item.exact_advance && item.text == " " {
                        word_space_extra
                    } else {
                        0.0
                    };
                emit_text_segment(
                    prims,
                    &item.text,
                    item.fmt,
                    item.pm_start,
                    item.pm_end,
                    pen_x,
                    baseline,
                    paint_width,
                    word_space_px.clone(),
                    item.level,
                    item.logical_order,
                    item.source_start,
                    item.source_end,
                    item.exact_advance,
                    line.synthetic_fallback,
                    geom.line_top,
                    line_bottom,
                    block_ref,
                    ctx.shape,
                    item.field,
                );
                if item.pm_start.is_some() {
                    emitted_positioned_text = true;
                }
                pen_x += paint_width;
            }
            LinePaintItem::Tab {
                run,
                width,
                logical_order,
                ..
            } => {
                let leader = run
                    .leader_glyphs
                    .as_ref()
                    .and_then(|leader| leader.glyph.as_deref())
                    .map(|_| "shaped".to_string())
                    .or_else(|| tab_leader_for(attrs, pen_x - geom.frag_x));
                if let Some(leader) = leader {
                    emit_tab_leader(
                        prims,
                        run,
                        &leader,
                        pen_x,
                        baseline,
                        *width,
                        *logical_order,
                        block_ref,
                    );
                    if run.pm_start.is_some() {
                        emitted_positioned_text = true;
                    }
                }
                pen_x += width;
            }
            LinePaintItem::Image {
                run: imr,
                pm_start,
                pm_end,
                single_image_line,
                level,
                logical_order,
            } => {
                // Image-only lines center; inline images sit on the baseline.
                let layout_width = imr
                    .rotation_bounds
                    .as_ref()
                    .and_then(|bounds| bounds.width)
                    .unwrap_or(imr.width);
                let layout_height = imr
                    .rotation_bounds
                    .as_ref()
                    .and_then(|bounds| bounds.height)
                    .unwrap_or(imr.height);
                let y = if *single_image_line {
                    geom.line_top + ((line.line_height - layout_height) / 2.0).max(0.0)
                } else {
                    baseline - layout_height
                };
                let rot = imr
                    .rotation_deg
                    .unwrap_or_else(|| rotation_degrees(imr.transform.as_deref()));
                let mut img_attrs = block_ref.attrs();
                img_attrs.doc_start = *pm_start;
                img_attrs.doc_end = *pm_end;
                img_attrs.logical_order = *logical_order;
                img_attrs.bidi_level = logical_order.map(|_| *level);
                stamp_image_run_attrs(&mut img_attrs, imr, pen_x, y);
                prims.push(Primitive::Image(ImagePrimitive {
                    rel_id: imr.src.clone(),
                    x: px(pen_x),
                    y: px(y),
                    w: px(layout_width),
                    h: px(layout_height),
                    rotation_deg: if rot != 0.0 { Some(px(rot)) } else { None },
                    opacity: imr.opacity.map(px),
                    filter: None,
                    decorative: imr.decorative.unwrap_or(false),
                    crop: crop_of(imr),
                    alt_text: capped_alt_text(imr.alt.as_deref()),
                    attrs: img_attrs,
                }));
                if imr.display_mode.as_deref() != Some("block")
                    && imr.wrap_type.as_deref() != Some("topAndBottom")
                {
                    pen_x += layout_width;
                }
            }
            LinePaintItem::LineBreak { pm_start, .. } => {
                if line_break_pos.is_none() {
                    line_break_pos = *pm_start;
                }
            }
        }
    }

    // A line with no positioned text still needs a doc position for hit
    // testing. A blank row from a line break carries the break's own inline
    // position; an empty paragraph line carries the paragraph's CONTENT
    // position, which is `pmStart + 1`. The fragment's own pmStart is the
    // paragraph boundary, and a caret cannot sit there — emitting it would land
    // every click and vertical move BEFORE the paragraph.
    if !emitted_positioned_text {
        let pos = line_break_pos.or(geom.frag_pm_start.map(|p| p + 1));
        if let Some(p) = pos {
            let mut marker_attrs = block_ref.attrs();
            marker_attrs.doc_start = Some(p);
            marker_attrs.doc_end = Some(p);
            prims.push(Primitive::Text(TextRunPrimitive {
                text: String::new(),
                x: px(geom.frag_x + pad_left + text_indent + left_offset),
                baseline_y: px(baseline),
                width: px(0.0),
                paint_clip: None,
                font: css_font(&RunFormattingIn::default()),
                color: "#000000".to_string(),
                letter_spacing: None,
                word_spacing: None,
                rtl: None,
                opacity: None,
                rotation_deg: None,
                horizontal_scale: None,
                all_caps: false,
                small_caps: false,
                hidden: false,
                text_shadow: None,
                text_outline: false,
                emphasis_mark: None,
                text_effect: None,
                attrs: marker_attrs,
            }));
        }
    }

    Some(LinePaintMetrics {
        end_x: pen_x,
        baseline,
    })
}

fn tab_leader_for(attrs: Option<&ParaAttrsIn>, current_x_px: f64) -> Option<String> {
    let stops = attrs?.tabs.as_ref()?;
    stops
        .iter()
        .filter(|stop| stop.val.as_deref() != Some("clear"))
        .filter_map(|stop| {
            let pos_px = stop.pos? / 15.0;
            if pos_px <= current_x_px + 0.5 {
                return None;
            }
            let leader = stop.leader.as_deref()?;
            if leader == "none" {
                return None;
            }
            Some((pos_px, leader.to_string()))
        })
        .min_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal))
        .map(|(_, leader)| leader)
}

fn emit_tab_leader(
    prims: &mut Vec<Primitive>,
    tab: &TabRunIn,
    leader: &str,
    x: f64,
    baseline: f64,
    width: f64,
    logical_order: Option<u64>,
    block_ref: &BlockRef,
) {
    if width <= 0.0 {
        return;
    }
    if let Some(measured) = &tab.leader_glyphs
        && let (Some(glyph), Some(advance)) = (measured.glyph.as_deref(), measured.advance)
        && !glyph.is_empty()
        && advance.is_finite()
        && advance > 0.0
    {
        let count = measured
            .count
            .unwrap_or_else(|| (width / advance).floor().max(1.0) as u64)
            .min(10_000);
        let mut fmt = tab.fmt.clone();
        if let Some(font) = &measured.font {
            fmt.font_family = Some(font.clone());
        }
        if let Some(size) = measured.font_size {
            fmt.font_size = Some(size);
        }
        if let Some(color) = &measured.color {
            fmt.color = Some(color.clone());
        }
        let mut attrs = block_ref.attrs();
        attrs.doc_start = tab.pm_start;
        attrs.doc_end = tab.pm_end;
        attrs.logical_order = logical_order.or(tab.fmt.logical_order);
        attrs.bidi_level = tab.fmt.bidi_level;
        attrs.leader_glyphs = Some(LeaderGlyphMetadata {
            glyph: Some(glyph.to_string()),
            count: Some(count),
            x: Some(px(x)),
            baseline_y: Some(px(baseline)),
            advance: Some(px(advance)),
            width: Some(px(width)),
            font: measured.font.clone(),
            font_id: None,
            size: measured.font_size.map(|size| px(size * 96.0 / 72.0)),
            color: Some(run_color(&fmt)),
            rtl: tab.fmt.rtl.filter(|rtl| *rtl),
        });
        prims.push(Primitive::Text(TextRunPrimitive {
            text: glyph.repeat(count as usize),
            x: px(x),
            baseline_y: px(baseline),
            width: px(width),
            paint_clip: None,
            font: css_font(&fmt),
            color: run_color(&fmt),
            letter_spacing: None,
            word_spacing: None,
            rtl: tab.fmt.rtl.filter(|rtl| *rtl),
            opacity: None,
            rotation_deg: None,
            horizontal_scale: horizontal_scale_of(&fmt),
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs,
        }));
        return;
    }
    let font_px = effective_font_px_of(&tab.fmt);
    let thickness = (font_px / 16.0).round().max(1.0);
    let (dashed, dotted, y, h) = match leader {
        "dot" | "middleDot" => (false, true, baseline - (font_px * 0.25).round(), thickness),
        "hyphen" => (true, false, baseline - (font_px * 0.25).round(), thickness),
        "underscore" => (
            false,
            false,
            baseline + (font_px * 0.1875).round(),
            thickness,
        ),
        "heavy" => (
            false,
            false,
            baseline + (font_px * 0.1875).round(),
            thickness.max(2.0),
        ),
        _ => return,
    };
    let mut attrs = block_ref.attrs();
    attrs.doc_start = tab.pm_start;
    attrs.doc_end = tab.pm_end;
    attrs.logical_order = logical_order.or(tab.fmt.logical_order);
    attrs.bidi_level = tab.fmt.bidi_level;
    prims.push(Primitive::Decoration(DecorationPrimitive {
        deco: DecoKind::Underline,
        x: px(x),
        y: px(y),
        w: px(width),
        h: px(h),
        color: run_color(&tab.fmt),
        dashed,
        dotted,
        attrs,
    }));
}

/// one styled text slice: comment tint and highlight paint behind the glyphs,
/// the text primitive itself, then underline/strike over it
#[allow(clippy::too_many_arguments)]
fn emit_text_segment(
    prims: &mut Vec<Primitive>,
    text: &str,
    fmt: &RunFormattingIn,
    pm_start: Option<i64>,
    pm_end: Option<i64>,
    x: f64,
    baseline: f64,
    width: f64,
    word_spacing: Option<Number>,
    bidi_level: u8,
    logical_order: Option<u64>,
    source_start: usize,
    source_end: usize,
    exact_advance: bool,
    synthetic_fallback: bool,
    line_top: f64,
    line_bottom: f64,
    block_ref: &BlockRef,
    shape: Option<&ShapeFonts<'_>>,
    field: Option<&FieldRunIn>,
) {
    let font_px = effective_font_px_of(fmt);
    let paint_baseline = baseline_y_of(fmt, baseline);
    let comment_ids: Option<Vec<String>> = fmt
        .comment_ids
        .as_ref()
        .filter(|ids| !ids.is_empty())
        .map(|ids| ids.iter().map(|id| id.to_string()).collect());
    let revision = if fmt.is_insertion == Some(true) || fmt.is_deletion == Some(true) {
        Some(Revision {
            author: fmt.change_author.clone().unwrap_or_default(),
            date: fmt.change_date.clone().unwrap_or_default(),
            revision_id: fmt
                .change_revision_id
                .map(|id| id.to_string())
                .unwrap_or_default(),
            kind: if fmt.is_insertion == Some(true) {
                RevisionKind::Ins
            } else {
                RevisionKind::Del
            },
        })
    } else {
        None
    };
    let mut attrs = block_ref.attrs();
    attrs.doc_start = pm_start;
    attrs.doc_end = pm_end;
    attrs.comment_ids = comment_ids.clone();
    attrs.revision = revision;
    attrs.href = hyperlink_href(fmt);
    attrs.inline_sdt_widget = fmt.inline_sdt_widget.clone();
    attrs.logical_order = logical_order.or(fmt.logical_order);
    attrs.bidi_level = exact_advance.then_some(bidi_level).or(fmt.bidi_level);
    attrs.lang = fmt.language.as_ref().and_then(|language| {
        if ooxml_text::level_is_rtl(bidi_level) {
            language.bidi.clone().or_else(|| language.latin.clone())
        } else {
            language
                .east_asia
                .clone()
                .or_else(|| language.latin.clone())
        }
    });
    if let Some(link) = &fmt.hyperlink {
        attrs.tooltip = link.tooltip.clone();
        attrs.link_title = link.tooltip.clone();
        attrs.link_target = link.target.clone();
        attrs.link_history = link.history;
        attrs.link_doc_location = link.doc_location.clone();
    }
    // inert field identity: type/instruction ride on the result primitives so
    // the a11y mirror can announce what the field is. Announce-only — nothing
    // downstream parses or executes the instruction.
    if let Some(field_run) = field {
        attrs.field = Some(field_metadata(field_run));
    }
    // footnote/endnote body reference mark → note_ref, the backlink hook
    // (the mirror renders it as a doc-noteref link to `oox-<kind>-<id>`)
    if let Some(id) = fmt.footnote_ref_id {
        attrs.note_ref = Some(NoteRefMetadata {
            kind: Some("footnote".to_string()),
            id: Some(id),
        });
    } else if let Some(id) = fmt.endnote_ref_id {
        attrs.note_ref = Some(NoteRefMetadata {
            kind: Some("endnote".to_string()),
            id: Some(id),
        });
    }

    // Highlight is the run font box, never the containing line band. Exact
    // slices carry their authoritative source boundaries for mirror/hit logic.
    if let Some(hl) = &fmt.highlight {
        let ascent = font_px * 0.8;
        let descent = font_px * 0.2;
        let mut highlight_attrs = attrs.clone();
        highlight_attrs.highlight_slice = Some(HighlightSliceMetadata {
            source_start: exact_advance.then_some(source_start as u64),
            source_end: exact_advance.then_some(source_end as u64),
            ascent: Some(px(ascent)),
            descent: Some(px(descent)),
            includes_trailing_whitespace: Some(
                text.chars().next_back().is_some_and(char::is_whitespace),
            ),
        });
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Highlight,
            x: px(x),
            y: px(paint_baseline - ascent),
            w: px(width),
            h: px(ascent + descent),
            color: hl.clone(),
            dashed: false,
            dotted: false,
            attrs: highlight_attrs,
        }));
    }
    // Comment ranges tint behind the glyphs.
    if comment_ids.is_some() {
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::CommentRange,
            x: px(x),
            y: px(line_top),
            w: px(width),
            h: px(line_bottom - line_top),
            color: "rgba(255, 212, 0, 0.15)".to_string(),
            dashed: false,
            dotted: false,
            attrs: attrs.clone(),
        }));
    }
    // Paint the insertion wash across the full line band.
    if fmt.is_insertion == Some(true) {
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Highlight,
            x: px(x),
            y: px(line_top),
            w: px(width),
            h: px(line_bottom - line_top),
            color: "rgba(52, 168, 83, 0.08)".to_string(),
            dashed: false,
            dotted: false,
            attrs: attrs.clone(),
        }));
    }
    // Suggested deletion: a red wash behind the glyphs. The red text itself
    // comes from run_color and the strike from the Strike decoration below.
    if fmt.is_deletion == Some(true) {
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Highlight,
            x: px(x),
            y: px(line_top),
            w: px(width),
            h: px(line_bottom - line_top),
            color: "rgba(211, 47, 47, 0.08)".to_string(),
            dashed: false,
            dotted: false,
            attrs: attrs.clone(),
        }));
    }

    let color = run_color(fmt);
    let paint_clip = synthetic_fallback.then(|| ClipRect {
        x: Some(px(x)),
        y: None,
        w: Some(px(width)),
        h: None,
    });
    // Unresolved or failed shaping falls back to a text primitive.
    let emitted_glyphs = match shape {
        Some(sf) => try_emit_glyph_runs(
            prims,
            sf,
            text,
            fmt,
            pm_start,
            pm_end,
            x,
            paint_baseline,
            &word_spacing,
            bidi_level,
            exact_advance.then_some(width),
            logical_order,
            &attrs,
            &color,
            &paint_clip,
        ),
        None => false,
    };
    if !emitted_glyphs {
        let mut text_attrs = attrs.clone();
        text_attrs.modern_effects = fmt.modern_effects.clone();
        prims.push(Primitive::Text(TextRunPrimitive {
            text: text.to_string(),
            x: px(x),
            baseline_y: px(paint_baseline),
            width: px(width),
            paint_clip,
            font: css_font(fmt),
            color: color.clone(),
            letter_spacing: fmt.letter_spacing.map(px),
            word_spacing,
            rtl: if ooxml_text::level_is_rtl(bidi_level) {
                Some(true)
            } else {
                None
            },
            opacity: None,
            rotation_deg: None,
            horizontal_scale: horizontal_scale_of(fmt),
            all_caps: fmt.all_caps == Some(true),
            small_caps: fmt.small_caps == Some(true),
            hidden: fmt.hidden == Some(true),
            text_shadow: text_shadow_of(fmt),
            text_outline: fmt.text_outline == Some(true),
            emphasis_mark: fmt.emphasis_mark.clone(),
            text_effect: fmt.text_effect.clone(),
            attrs: text_attrs,
        }));
    }

    // decoration thickness/offsets derived from the font size (deterministic
    // stand-ins for the browser's UA decoration metrics)
    let thickness = (font_px / 16.0).round().max(1.0);
    let has_underline = matches!(
        &fmt.underline,
        Some(Value::Bool(true)) | Some(Value::Object(_))
    );
    if has_underline {
        let (deco_color, underline_style) = match &fmt.underline {
            Some(Value::Object(o)) => (
                o.get("color")
                    .and_then(|c| c.as_str())
                    .map(|s| s.to_string())
                    .unwrap_or_else(|| color.clone()),
                display_border_style(o.get("style").and_then(Value::as_str)),
            ),
            _ => (color.clone(), DisplayBorderStyle::Solid),
        };
        let mut underline_attrs = attrs.clone();
        underline_attrs.style = Some(underline_style);
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Underline,
            x: px(x),
            y: px(paint_baseline + (font_px * 0.1875).round()),
            w: px(width),
            h: px(thickness),
            color: deco_color,
            dashed: matches!(
                underline_style,
                DisplayBorderStyle::Dashed
                    | DisplayBorderStyle::DashDot
                    | DisplayBorderStyle::DashDotDot
            ),
            dotted: underline_style == DisplayBorderStyle::Dotted,
            attrs: underline_attrs,
        }));
    } else if fmt.is_insertion == Some(true) {
        // Suggested insertion: a green dashed rule under the run, reusing the
        // offset and thickness of an explicit underline; the dashed flag drives
        // the canvas
        // rule. `else if` so an already-underlined inserted run keeps a single
        // rule rather than stacking two lines at the same baseline offset.
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Underline,
            x: px(x),
            y: px(paint_baseline + (font_px * 0.1875).round()),
            w: px(width),
            h: px(thickness),
            color: "#2e7d32".to_string(),
            dashed: true,
            dotted: false,
            attrs: attrs.clone(),
        }));
    } else if fmt.hidden == Some(true) {
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Underline,
            x: px(x),
            y: px(paint_baseline + (font_px * 0.1875).round()),
            w: px(width),
            h: px(thickness),
            color: color.clone(),
            dashed: false,
            dotted: true,
            attrs: attrs.clone(),
        }));
    }
    // Deletions strike through in the revision colour.
    if fmt.strike == Some(true) || fmt.is_deletion == Some(true) {
        prims.push(Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Strike,
            x: px(x),
            y: px(paint_baseline - (font_px * 0.3).round()),
            w: px(width),
            h: px(thickness),
            color,
            dashed: false,
            dotted: false,
            attrs,
        }));
    }
}

/// Shape one styled text slice into [`GlyphRunPrimitive`]s and push them, or
/// return `false` to signal the caller to emit a [`TextRunPrimitive`]
/// instead. Returns `false` (touching nothing) when the run has no text, its
/// font family has no chain, or any subrange fails to shape — so the fallback
/// is always a clean whole-segment TextRunPrimitive, never a partial mix.
///
/// A run can span several fallback fonts, but a GlyphRun is single-font by
/// contract, so the segment is split into maximal same-font subranges (the same
/// policy as ooxml-text `measure_plain_text`) and one GlyphRun is emitted per
/// subrange. RTL subranges are placed in visual order. The pen advance
/// accumulates ACROSS subranges so contiguous glyphs keep flowing; each glyph's
/// `x` is `x + accumulated_advance + x_offset`, its `y` is `baseline - y_offset`
/// (shaped `y_offset` is up, screen y is down).
/// Justification stretch (`word_spacing`) is folded into the advance after each
/// U+0020 cluster, matching how `emit_line` computed the line width.
#[allow(clippy::too_many_arguments)]
fn try_emit_glyph_runs(
    prims: &mut Vec<Primitive>,
    shape_fonts: &ShapeFonts<'_>,
    text: &str,
    fmt: &RunFormattingIn,
    pm_start: Option<i64>,
    pm_end: Option<i64>,
    x: f64,
    baseline: f64,
    word_spacing: &Option<Number>,
    bidi_level: u8,
    exact_width: Option<f64>,
    logical_order: Option<u64>,
    attrs: &DocAttrs,
    color: &str,
    paint_clip: &Option<ClipRect>,
) -> bool {
    if text.is_empty() {
        return false;
    }
    // Browser-shaped text primitives are still the faithful path for CSS-only
    // effects such as small-caps, text-emphasis, and hidden dotted underlines.
    if text_primitive_requires_browser_path(fmt) {
        return false;
    }
    let family = fmt.font_family.as_deref().unwrap_or(DEFAULT_FONT_FAMILY);
    let bold = fmt.bold == Some(true);
    let italic = fmt.italic == Some(true);
    let Some(chain) = shape_fonts.chain_for(family, bold, italic) else {
        return false;
    };
    if chain.is_empty() {
        return false;
    }

    let size_px = effective_font_px_of(fmt);
    let ws_px = word_spacing.as_ref().map(num_f64).unwrap_or(0.0);
    let direction = shape_direction_for_level(bidi_level);
    let rtl = if direction == ooxml_text::ShapeDirection::Rtl {
        Some(true)
    } else {
        None
    };

    // resolve each char to a font (chain head fallback per ooxml-text). A run
    // covered by a single font — the overwhelmingly common case — is detected
    // in this validation pass without any per-char scratch allocation.
    let mut first_font: Option<ooxml_text::FontId> = None;
    let mut uniform = true;
    for ch in text.chars() {
        match shape_fonts.resolve(chain, ch) {
            Some(font) => match first_font {
                None => first_font = Some(font),
                Some(first) if font != first => uniform = false,
                Some(_) => {}
            },
            None => return false,
        }
    }
    let Some(first_font) = first_font else {
        return false;
    };

    /// one maximal same-font subrange: its owned text slice plus the utf16
    /// offsets of its char boundaries within the whole segment
    struct ShapeSegment {
        text: String,
        font: ooxml_text::FontId,
        utf16_start: usize,
        utf16_end: usize,
    }

    let mut segments: Vec<ShapeSegment> = if uniform {
        vec![ShapeSegment {
            text: text.to_owned(),
            font: first_font,
            utf16_start: 0,
            // the single segment is also the last one, which closes on pm_end
            // without consulting utf16_end
            utf16_end: 0,
        }]
    } else {
        let chars: Vec<char> = text.chars().collect();
        let mut fonts: Vec<ooxml_text::FontId> = Vec::with_capacity(chars.len());
        for &ch in &chars {
            match shape_fonts.resolve(chain, ch) {
                Some(font) => fonts.push(font),
                None => return false,
            }
        }
        // utf16 prefix table: prefix[i] = utf16 length of chars[..i], so a
        // subrange's doc span is two lookups instead of two O(n) rescans
        let mut prefix = Vec::with_capacity(chars.len() + 1);
        let mut units = 0usize;
        prefix.push(0);
        for &ch in &chars {
            units += ch.len_utf16();
            prefix.push(units);
        }
        let mut segments = Vec::new();
        let mut ci = 0usize;
        while ci < chars.len() {
            let font = fonts[ci];
            let mut cj = ci + 1;
            while cj < chars.len() && fonts[cj] == font {
                cj += 1;
            }
            segments.push(ShapeSegment {
                text: chars[ci..cj].iter().collect(),
                font,
                utf16_start: prefix[ci],
                utf16_end: prefix[cj],
            });
            ci = cj;
        }
        segments
    };

    // build the glyph runs into a local buffer; commit only when every subrange
    // shapes, so a mid-segment failure never leaves a partial run behind
    let mut local: Vec<Primitive> = Vec::new();
    let mut acc = 0.0_f64; // pen advance from the segment origin `x`
    let segment_count = segments.len();
    let range_order: Box<dyn Iterator<Item = usize>> =
        if direction == ooxml_text::ShapeDirection::Rtl {
            Box::new((0..segment_count).rev())
        } else {
            Box::new(0..segment_count)
        };
    for range_index in range_order {
        let font = segments[range_index].font;
        // each index is visited exactly once, so the segment text moves into
        // its primitive instead of cloning
        let sub_text = std::mem::take(&mut segments[range_index].text);
        let (sub_start_utf16, sub_end_utf16) = (
            segments[range_index].utf16_start,
            segments[range_index].utf16_end,
        );
        let is_last = range_index + 1 == segment_count;
        let glyphs = match ooxml_text::shape_with_direction(
            shape_fonts.store,
            font,
            &sub_text,
            size_px as f32,
            &[],
            direction,
        ) {
            Ok(g) => g,
            Err(_) => return false,
        };

        let sub_bytes = sub_text.as_bytes();
        let mut placed: Vec<PlacedGlyph> = Vec::with_capacity(glyphs.len());
        for g in &glyphs {
            // Fold the justified U+0020 stretch into the glyph advance.
            let mut advance = g.x_advance as f64;
            if ws_px != 0.0 && sub_bytes.get(g.cluster as usize) == Some(&b' ') {
                advance += ws_px;
            }
            placed.push(PlacedGlyph {
                id: g.glyph_id,
                x: round3(x + acc + g.x_offset as f64),
                y: round3(baseline - g.y_offset as f64),
                cluster: g.cluster,
                advance: round3(advance),
                logical_order,
                bidi_level: exact_width.map(|_| bidi_level),
            });
            acc += advance;
        }

        // text runs are 1 document position per char, so a subrange's doc span shifts
        // pm_start by its char offset; the last subrange closes on pm_end so a
        // single-subrange run carries exactly the segment's [pm_start, pm_end]
        let mut sub_attrs = attrs.clone();
        sub_attrs.doc_start = pm_start.map(|p| p + sub_start_utf16 as i64);
        sub_attrs.doc_end = if is_last {
            pm_end
        } else {
            pm_start.map(|p| p + sub_end_utf16 as i64)
        };
        // the resolved CSS face for the canvas fillText safety net (glyph
        // outlines unavailable) — same shorthand the TextRunPrimitive would
        // carry, so the fallback keeps family/weight/style
        sub_attrs.fallback_font = Some(css_font(fmt));
        sub_attrs.modern_effects = fmt.modern_effects.clone();

        local.push(Primitive::GlyphRun(GlyphRunPrimitive {
            font_id: font.to_u32(),
            size: round3(size_px),
            color: color.to_string(),
            text: sub_text,
            glyphs: placed,
            paint_clip: paint_clip.clone(),
            word_spacing: word_spacing.clone(),
            rtl,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: horizontal_scale_of(fmt),
            all_caps: fmt.all_caps == Some(true),
            small_caps: fmt.small_caps == Some(true),
            hidden: fmt.hidden == Some(true),
            text_shadow: text_shadow_of(fmt),
            text_outline: fmt.text_outline == Some(true),
            emphasis_mark: fmt.emphasis_mark.clone(),
            text_effect: fmt.text_effect.clone(),
            attrs: sub_attrs,
        }));
    }

    if let Some(target) = exact_width
        && let Some(Primitive::GlyphRun(run)) = local.last_mut()
        && let Some(last) = run.glyphs.last_mut()
    {
        last.advance = round3(last.advance + target - acc);
    }

    prims.extend(local);
    true
}

/// field instruction -> display text (PAGE/NUMPAGES from the page context;
/// DATE/TIME deliberately resolve to the stored fallback for determinism)
fn field_text(f: &FieldRunIn, ctx: &RenderCtx<'_>) -> String {
    match f.field_type.as_deref() {
        Some("PAGE") => ctx.page_number.to_string(),
        Some("NUMPAGES") => ctx.total_pages.to_string(),
        _ => f.fallback.clone().unwrap_or_default(),
    }
}

/// field instruction cap — announcement identity, not a full field-code view
pub const MAX_FIELD_INSTRUCTION_CHARS: usize = 1024;

/// Inert accessibility identity of a field run: its resolved category, raw type
/// token, and bounded instruction. The instruction is file-derived and
/// attacker-controlled; it is carried for announcement ONLY — nothing here or
/// downstream evaluates it (field codes render inert; see the repo security guidelines).
fn field_metadata(f: &FieldRunIn) -> FieldMetadata {
    FieldMetadata {
        category: f.field_type.clone(),
        r#type: capped_string(f.raw_type.as_deref(), MAX_FIELD_INSTRUCTION_CHARS),
        instruction: capped_string(f.instruction.as_deref(), MAX_FIELD_INSTRUCTION_CHARS),
    }
}

fn border_visible(b: &BorderEdgeIn) -> bool {
    !matches!(b.style.as_deref(), Some("none") | Some("nil")) && b.width != Some(0.0)
}

fn display_border_style(style: Option<&str>) -> DisplayBorderStyle {
    match style.unwrap_or("single") {
        "double" => DisplayBorderStyle::Double,
        "dotted" => DisplayBorderStyle::Dotted,
        "dash" | "dashed" | "dashSmallGap" | "dashLong" | "dashLargeGap" => {
            DisplayBorderStyle::Dashed
        }
        "dotDash" | "dashDot" => DisplayBorderStyle::DashDot,
        "dotDotDash" | "dashDotDot" => DisplayBorderStyle::DashDotDot,
        "triple" => DisplayBorderStyle::Triple,
        "thinThickSmallGap" | "thinThickMediumGap" | "thinThickLargeGap" | "thinThick" => {
            DisplayBorderStyle::ThinThick
        }
        "thickThinSmallGap" | "thickThinMediumGap" | "thickThinLargeGap" | "thickThin" => {
            DisplayBorderStyle::ThickThin
        }
        "wave" => DisplayBorderStyle::Wave,
        "doubleWave" => DisplayBorderStyle::DoubleWave,
        "threeDEngrave" | "groove" => DisplayBorderStyle::Groove,
        "threeDEmboss" | "ridge" => DisplayBorderStyle::Ridge,
        "inset" => DisplayBorderStyle::Inset,
        "outset" => DisplayBorderStyle::Outset,
        _ => DisplayBorderStyle::Solid,
    }
}

fn border_dash(style: DisplayBorderStyle, width: f64) -> Option<Vec<Number>> {
    let unit = width.max(1.0);
    match style {
        DisplayBorderStyle::Dotted => Some(vec![px(unit), px(unit * 1.5)]),
        DisplayBorderStyle::Dashed => Some(vec![px(unit * 4.0), px(unit * 2.0)]),
        DisplayBorderStyle::DashDot => Some(vec![
            px(unit * 4.0),
            px(unit * 1.5),
            px(unit),
            px(unit * 1.5),
        ]),
        DisplayBorderStyle::DashDotDot => Some(vec![
            px(unit * 4.0),
            px(unit * 1.5),
            px(unit),
            px(unit * 1.5),
            px(unit),
            px(unit * 1.5),
        ]),
        _ => None,
    }
}

fn borders_edge_equal(a: Option<&BorderEdgeIn>, b: Option<&BorderEdgeIn>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => a.style == b.style && a.width == b.width && a.color == b.color,
        _ => false,
    }
}

/// ECMA-376 §17.3.1.24 border grouping: adjacent paragraphs with identical
/// border definitions form a group
fn borders_form_group(a: Option<&ParaBordersIn>, b: Option<&ParaBordersIn>) -> bool {
    let (Some(a), Some(b)) = (a, b) else {
        return false;
    };
    borders_edge_equal(a.top.as_ref(), b.top.as_ref())
        && borders_edge_equal(a.bottom.as_ref(), b.bottom.as_ref())
        && borders_edge_equal(a.left.as_ref(), b.left.as_ref())
        && borders_edge_equal(a.right.as_ref(), b.right.as_ref())
        && borders_edge_equal(a.between.as_ref(), b.between.as_ref())
}

fn points_to_px(points: f64) -> f64 {
    points * 96.0 / 72.0
}

fn eighths_to_px(eighths: f64) -> f64 {
    points_to_px(eighths / 8.0)
}

fn page_border_should_render(page_number: u64, display: Option<&str>) -> bool {
    match display.unwrap_or("allPages") {
        "firstPage" => page_number == 1,
        "notFirstPage" => page_number != 1,
        _ => true,
    }
}

fn page_border_visible(border: Option<&PageBorderSpecIn>) -> bool {
    let Some(border) = border else {
        return false;
    };
    !matches!(border.style.as_deref(), Some("none") | Some("nil"))
}

fn page_border_space_px(border: Option<&PageBorderSpecIn>) -> f64 {
    border
        .and_then(|b| b.space)
        .map(points_to_px)
        .unwrap_or(0.0)
}

fn page_border_style(style: Option<&str>) -> &'static str {
    match style.unwrap_or("single") {
        "double" | "triple" => "double",
        "dotted" => "dotted",
        "dashed" | "dashSmallGap" => "dashed",
        "threeDEmboss" => "ridge",
        "threeDEngrave" => "groove",
        "outset" => "outset",
        "inset" => "inset",
        _ => "solid",
    }
}

fn normalize_hex_color(raw: &str) -> String {
    if raw.starts_with('#') || raw.starts_with("rgb") || raw == "transparent" {
        return raw.to_string();
    }
    if raw.eq_ignore_ascii_case("auto") {
        return "#000000".to_string();
    }
    if raw.len() == 6 && raw.chars().all(|c| c.is_ascii_hexdigit()) {
        return format!("#{raw}");
    }
    raw.to_string()
}

fn page_border_color(color: Option<&Value>) -> String {
    let Some(color) = color else {
        return "#000000".to_string();
    };
    if let Some(s) = color.as_str() {
        return normalize_hex_color(s);
    }
    if color.get("auto").and_then(Value::as_bool).unwrap_or(false) {
        return "#000000".to_string();
    }
    if let Some(rgb) = color.get("rgb").and_then(Value::as_str) {
        return normalize_hex_color(rgb);
    }
    "#000000".to_string()
}

fn page_border_side(border: Option<&PageBorderSpecIn>) -> Option<PageBorderSide> {
    if !page_border_visible(border) {
        return None;
    }
    let border = border?;
    let mut width = eighths_to_px(border.size.unwrap_or(6.0)).max(1.0);
    let style = page_border_style(border.style.as_deref()).to_string();
    if style == "double" && width < 3.0 {
        width = 3.0;
    }
    Some(PageBorderSide {
        width: px(width),
        color: page_border_color(border.color.as_ref()),
        style,
    })
}

fn page_border_primitive(
    options: &RenderOptionsIn,
    page: &PageIn,
    page_number: u64,
) -> Option<PageBorderPrimitive> {
    let pb = options.page_borders.as_ref()?;
    if !page_border_should_render(page_number, pb.display.as_deref()) {
        return None;
    }
    if ![
        pb.top.as_ref(),
        pb.right.as_ref(),
        pb.bottom.as_ref(),
        pb.left.as_ref(),
    ]
    .into_iter()
    .any(page_border_visible)
    {
        return None;
    }

    let top_offset = page_border_space_px(pb.top.as_ref());
    let right_offset = page_border_space_px(pb.right.as_ref());
    let bottom_offset = page_border_space_px(pb.bottom.as_ref());
    let left_offset = page_border_space_px(pb.left.as_ref());
    let (top, right, bottom, left) = if pb.offset_from.as_deref() == Some("page") {
        (top_offset, right_offset, bottom_offset, left_offset)
    } else {
        (
            (page.margins.top - top_offset).max(0.0),
            (page.margins.right - right_offset).max(0.0),
            (page.margins.bottom - bottom_offset).max(0.0),
            (page.margins.left - left_offset).max(0.0),
        )
    };
    let w = (page.size.w - left - right).max(0.0);
    let h = (page.size.h - top - bottom).max(0.0);
    Some(PageBorderPrimitive {
        x: px(left),
        y: px(top),
        w: px(w),
        h: px(h),
        z_order: if pb.z_order.as_deref() == Some("back") {
            Some(PageBorderZOrder::Back)
        } else {
            Some(PageBorderZOrder::Front)
        },
        top: page_border_side(pb.top.as_ref()),
        right: page_border_side(pb.right.as_ref()),
        bottom: page_border_side(pb.bottom.as_ref()),
        left: page_border_side(pb.left.as_ref()),
    })
}

/// paragraph borders as line primitives (role 'border'): top only when the
/// fragment opens its group (between-rule otherwise), bottom only when it
/// closes it, left/right on every member, plus the §17.3.1.4 bar edge
#[allow(clippy::too_many_arguments)]
fn emit_paragraph_borders(
    prims: &mut Vec<Primitive>,
    borders: &ParaBordersIn,
    prev: Option<&ParaBordersIn>,
    next: Option<&ParaBordersIn>,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    indent_left: f64,
    indent_right: f64,
) {
    let grouped_with_prev = borders_form_group(prev, Some(borders));
    let grouped_with_next = borders_form_group(Some(borders), next);
    let top = if grouped_with_prev {
        borders.between.as_ref()
    } else {
        borders.top.as_ref()
    };
    let bottom = if grouped_with_next {
        None
    } else {
        borders.bottom.as_ref()
    };

    let left_space = borders.left.as_ref().and_then(|b| b.space).unwrap_or(0.0);
    let right_space = borders.right.as_ref().and_then(|b| b.space).unwrap_or(0.0);
    let lx = x + indent_left - left_space;
    let rx = x + width - indent_right + right_space;
    let ty = y - top.and_then(|b| b.space).unwrap_or(0.0);
    let by = y + height + bottom.and_then(|b| b.space).unwrap_or(0.0);

    let mut push = |x1: f64, y1: f64, x2: f64, y2: f64, b: &BorderEdgeIn| {
        if !border_visible(b) {
            return;
        }
        let style = display_border_style(b.style.as_deref());
        let width = b.width.unwrap_or(1.0);
        prims.push(Primitive::Line(LinePrimitive {
            x1: px(x1),
            y1: px(y1),
            x2: px(x2),
            y2: px(y2),
            stroke_width: px(width),
            color: b.color.clone().unwrap_or_else(|| "#000000".to_string()),
            dash: border_dash(style, width),
            role: Some(LineRole::Border),
            border_style: Some(style),
            border_owner: Some(BorderOwner::Paragraph),
            ..LinePrimitive::contract_defaults()
        }));
    };

    if let Some(b) = top {
        push(lx, ty, rx, ty, b);
    }
    if let Some(b) = bottom {
        push(lx, by, rx, by, b);
    }
    if let Some(b) = borders.left.as_ref() {
        push(lx, ty, lx, by, b);
    }
    if let Some(b) = borders.right.as_ref() {
        push(rx, ty, rx, by, b);
    }
    // bar border: vertical decorative rule left of the text (§17.3.1.4)
    if let Some(b) = borders.bar.as_ref() {
        push(x - 8.0, y, x - 8.0, y + height, b);
    }
}

// ---------------------------------------------------------------------------
// floating images + text boxes
// ---------------------------------------------------------------------------

/// Page coordinate frame for anchored floats, in pixels.
struct PageFloatGeom {
    page_width: f64,
    page_height: f64,
    margin_left: f64,
    margin_top: f64,
    content_width: f64,
    content_height: f64,
}

/// an anchor band: `base` is the band's origin (content-relative px) and `size`
/// its extent; `size == 0` marks a zero-width band (align falls back to base/0).
struct AnchorBand {
    base: f64,
    size: f64,
}

/// Converts EMU to rounded pixels at 96 DPI.
fn emu_to_px(emu: f64) -> f64 {
    (emu * 96.0 / 914400.0).round()
}

/// Resolves the horizontal band for a `relativeFrom` value.
fn horizontal_anchor_band(relative_to: Option<&str>, geom: &PageFloatGeom) -> AnchorBand {
    match relative_to {
        Some("page") => AnchorBand {
            base: -geom.margin_left,
            size: geom.page_width,
        },
        Some("leftMargin") => AnchorBand {
            base: -geom.margin_left,
            size: geom.margin_left,
        },
        Some("rightMargin") => AnchorBand {
            base: geom.content_width,
            size: geom.margin_left,
        },
        Some("character") => AnchorBand {
            base: 0.0,
            size: 0.0,
        },
        // column / margin / insideMargin / outsideMargin / default
        _ => AnchorBand {
            base: 0.0,
            size: geom.content_width,
        },
    }
}

/// Resolves the vertical band for a `relativeFrom` value.
fn vertical_anchor_band(
    relative_to: Option<&str>,
    fragment_y: f64,
    geom: &PageFloatGeom,
) -> AnchorBand {
    match relative_to {
        Some("paragraph") | Some("line") => AnchorBand {
            base: fragment_y,
            size: 0.0,
        },
        Some("page") => AnchorBand {
            base: -geom.margin_top,
            size: geom.page_height,
        },
        Some("topMargin") => AnchorBand {
            base: -geom.margin_top,
            size: geom.margin_top,
        },
        Some("bottomMargin") => AnchorBand {
            base: geom.content_height,
            size: geom.margin_top,
        },
        // margin / insideMargin / outsideMargin / default
        _ => AnchorBand {
            base: 0.0,
            size: geom.content_height,
        },
    }
}

/// Resolves a floating image run's OOXML anchor to a content-relative origin.
///
/// Each axis reads its `relativeFrom` band, then applies either an explicit
/// offset or an alignment inside that band. `fragment_y` is the anchoring
/// paragraph fragment's content-relative top, which is the base for the
/// `paragraph` and `line` bands. Anchors are resolved here rather than read off
/// the fragment, because the layout does not store them.
fn resolve_anchored_position(
    imr: &ImageRunIn,
    fragment_y: f64,
    geom: &PageFloatGeom,
) -> (f64, f64) {
    let x = match imr.position.as_ref().and_then(|p| p.horizontal.as_ref()) {
        None => {
            if imr.css_float.as_deref() == Some("right") {
                geom.content_width - image_layout_width(imr)
            } else {
                0.0
            }
        }
        Some(h) => {
            let band = horizontal_anchor_band(h.relative_to.as_deref(), geom);
            match h.align.as_deref() {
                Some("right") => {
                    if band.size != 0.0 {
                        band.base + band.size - image_layout_width(imr)
                    } else {
                        0.0
                    }
                }
                Some("left") => band.base,
                Some("center") => {
                    if band.size != 0.0 {
                        band.base + (band.size - image_layout_width(imr)) / 2.0
                    } else {
                        0.0
                    }
                }
                _ => match h.pos_offset {
                    Some(off) => band.base + emu_to_px(off),
                    None => band.base,
                },
            }
        }
    };

    let y = match imr.position.as_ref().and_then(|p| p.vertical.as_ref()) {
        None => fragment_y,
        Some(v) => {
            let band = vertical_anchor_band(v.relative_to.as_deref(), fragment_y, geom);
            match v.align.as_deref() {
                Some("top") => band.base,
                Some("center") => {
                    if band.size != 0.0 {
                        band.base + (band.size - image_layout_height(imr)) / 2.0
                    } else {
                        fragment_y
                    }
                }
                Some("bottom") => {
                    if band.size != 0.0 {
                        band.base + band.size - image_layout_height(imr)
                    } else {
                        fragment_y
                    }
                }
                _ => match v.pos_offset {
                    Some(off) => band.base + emu_to_px(off),
                    None => {
                        if matches!(v.relative_to.as_deref(), Some("paragraph") | Some("line")) {
                            fragment_y
                        } else {
                            band.base
                        }
                    }
                },
            }
        }
    };

    (x, y)
}

/// Emits a paragraph's floating image runs at their resolved page rectangles.
/// `want_behind` selects the pass: behind-document floats paint before body
/// content, the rest after.
fn emit_paragraph_floating_images(
    prims: &mut Vec<Primitive>,
    block: &ParagraphBlockIn,
    frag_y: f64,
    geom: &PageFloatGeom,
    want_behind: bool,
) {
    let block_ref = BlockRef::of(&block.id);
    // Float anchors resolve against the content area, not the page.
    let fragment_content_y = frag_y - geom.margin_top;
    for run in &block.runs {
        let RunIn::Image(imr) = run else { continue };
        if !is_floating_image_run(imr) {
            continue;
        }
        let is_behind = imr.wrap_type.as_deref() == Some("behind");
        if is_behind != want_behind {
            continue;
        }
        let (x, y) = resolve_anchored_position(imr, fragment_content_y, geom);
        // Content-relative back to page-local.
        let page_x = geom.margin_left + x;
        let page_y = geom.margin_top + y;
        let rot = imr
            .rotation_deg
            .unwrap_or_else(|| rotation_degrees(imr.transform.as_deref()));
        let layout_width = image_layout_width(imr);
        let layout_height = image_layout_height(imr);
        let mut attrs = block_ref.attrs();
        attrs.doc_start = imr.pm_start;
        attrs.doc_end = imr.pm_end;
        stamp_image_run_attrs(&mut attrs, imr, page_x, page_y);
        attrs.sdt = sdt_attrs_from_groups(&block.sdt_groups);
        attrs.sdt_path = sdt_path_from_groups(&block.sdt_groups);
        prims.push(Primitive::Image(ImagePrimitive {
            rel_id: imr.src.clone(),
            x: px(page_x),
            y: px(page_y),
            w: px(layout_width),
            h: px(layout_height),
            rotation_deg: if rot != 0.0 { Some(px(rot)) } else { None },
            opacity: imr.opacity.map(px),
            filter: None,
            decorative: imr.decorative.unwrap_or(false),
            crop: crop_of(imr),
            alt_text: capped_alt_text(imr.alt.as_deref()),
            attrs,
        }));
    }
}

/// paint one DrawingML shape fragment: a page-placed path primitive carrying
/// the shape's scaled geometry, paint, transform, document range, and optional
/// inner paragraphs when their measurements are present.
fn emit_shape_fragment(
    prims: &mut Vec<Primitive>,
    frag: &ShapeFragmentIn,
    block: &ShapeBlockIn,
    ctx: &RenderCtx<'_>,
) {
    let stamp_from = prims.len();
    let block_ref = BlockRef::of(&frag.block_id);
    let mut attrs = block_ref.attrs();
    attrs.doc_start = frag
        .doc_start
        .or(frag.pm_start)
        .or(block.doc_start)
        .or(block.pm_start);
    attrs.doc_end = frag
        .doc_end
        .or(frag.pm_end)
        .or(block.doc_end)
        .or(block.pm_end);
    attrs.sdt = sdt_attrs_from_groups(&block.sdt_groups);
    attrs.sdt_path = sdt_path_from_groups(&block.sdt_groups);
    attrs.aria_label = block.title.clone();
    attrs.aria_description = block.description.clone();
    // decorative is carried by ShapePrimitive's own field; duplicating it on
    // the flattened attrs produced two identical JSON keys that serde_json's
    // Map used to collapse silently. One authoritative slot keeps the encoded
    // object duplicate-free.
    attrs.hidden_object = block.hidden.filter(|hidden| *hidden);
    attrs.logical_order = block.relative_height;
    attrs.group_id = block
        .scene
        .as_ref()
        .and_then(|scene| scene.pointer("/root/id"))
        .and_then(Value::as_str)
        .map(str::to_string);
    attrs.fill_paint = shape_fill_paint(block.fill.as_ref());
    attrs.stroke_paint = shape_stroke_paint(block.stroke.as_ref());
    attrs.effects = block.effects.clone();
    attrs.effect_extent = block.effect_extent.clone();
    attrs.drawing_scene = block.scene.clone();
    attrs.text_body_properties = block.text_body_properties.clone();

    let decorative = block.decorative.unwrap_or_else(|| {
        block.inner_text.is_empty() && block.title.is_none() && block.description.is_none()
    });

    prims.push(Primitive::Shape(ShapePrimitive {
        x: px(frag.x),
        y: px(frag.y),
        w: px(frag.width),
        h: px(frag.height),
        geometry_path: scale_shape_path(&block.geometry_path, frag),
        fill: shape_fill_color(block.fill.as_ref()),
        stroke: shape_stroke(block.stroke.as_ref()),
        transform: shape_transform(block.transform.as_ref()),
        decorative,
        attrs,
    }));

    if !block.inner_text.is_empty() {
        let body_margins = block
            .text_body_properties
            .as_ref()
            .and_then(|properties| properties.get("margins"));
        let margin_left = body_margins
            .and_then(|margins| margins.get("left"))
            .and_then(Value::as_f64)
            .unwrap_or(0.0);
        let margin_right = body_margins
            .and_then(|margins| margins.get("right"))
            .and_then(Value::as_f64)
            .unwrap_or(0.0);
        let margin_top = body_margins
            .and_then(|margins| margins.get("top"))
            .and_then(Value::as_f64)
            .unwrap_or(0.0);
        let margin_bottom = body_margins
            .and_then(|margins| margins.get("bottom"))
            .and_then(Value::as_f64)
            .unwrap_or(0.0);
        let content_height: f64 = block
            .inner_measures
            .iter()
            .map(|measure| measure.total_height)
            .sum();
        let available_height = (frag.height - margin_top - margin_bottom).max(0.0);
        let anchor = block
            .text_body_properties
            .as_ref()
            .and_then(|properties| properties.get("anchor"))
            .and_then(Value::as_str);
        let anchor_offset = match anchor {
            Some("middle") => ((available_height - content_height) / 2.0).max(0.0),
            Some("bottom") => (available_height - content_height).max(0.0),
            _ => 0.0,
        };
        let content_x = frag.x + margin_left;
        let content_width = (frag.width - margin_left - margin_right).max(0.0);
        let mut y_offset = margin_top + anchor_offset;
        for (i, para) in block.inner_text.iter().enumerate() {
            let Some(pm) = block.inner_measures.get(i) else {
                continue;
            };
            let para_y = frag.y + y_offset;
            let synthetic = ParagraphFragmentIn {
                block_id: para.id.clone(),
                x: content_x,
                y: para_y,
                width: content_width,
                height: pm.total_height,
                from_line: 0,
                to_line: pm.lines.len(),
                pm_start: para.pm_start,
                pm_end: para.pm_end,
                carried_from_prev: None,
                carried_to_next: None,
            };
            emit_paragraph_fragment(
                prims, &synthetic, para, pm, ctx, content_x, para_y, None, None, true, true,
            );
            y_offset += pm.total_height;
        }
    }

    for child in &block.children {
        let child_frag = ShapeFragmentIn {
            block_id: child.id.clone(),
            x: frag.x + child.x.unwrap_or(0.0),
            y: frag.y + child.y.unwrap_or(0.0),
            width: child.width,
            height: child.height,
            doc_start: child.doc_start.or(child.pm_start),
            doc_end: child.doc_end.or(child.pm_end),
            pm_start: child.pm_start,
            pm_end: child.pm_end,
            is_anchored: None,
            z_index: None,
        };
        emit_shape_fragment(prims, &child_frag, child, ctx);
    }

    stamp_sdt_range(&mut prims[stamp_from..], &block.sdt_groups, false);
}

fn shape_fill_color(fill: Option<&ShapeFillIn>) -> Option<String> {
    let fill = fill?;
    ooxml_drawingml::resolve_shape_fill_color(fill.kind.as_deref(), fill.color.as_deref())
}

fn shape_fill_paint(fill: Option<&ShapeFillIn>) -> Option<Value> {
    let fill = fill?;
    let mut paint = serde_json::Map::new();
    if let Some(kind) = &fill.kind {
        paint.insert("kind".to_string(), Value::String(kind.clone()));
    }
    if let Some(color) = &fill.color {
        paint.insert("color".to_string(), Value::String(color.clone()));
    }
    if let Some(angle) = fill.gradient_angle.and_then(Number::from_f64) {
        paint.insert("angle".to_string(), Value::Number(angle));
    }
    if !fill.gradient_stops.is_empty() {
        paint.insert(
            "stops".to_string(),
            Value::Array(fill.gradient_stops.clone()),
        );
    }
    for (key, value) in [
        ("gradientType", fill.gradient_type.as_ref()),
        ("patternPreset", fill.pattern_preset.as_ref()),
        ("foregroundColor", fill.foreground_color.as_ref()),
        ("backgroundColor", fill.background_color.as_ref()),
        ("pictureRelId", fill.picture_rel_id.as_ref()),
        ("pictureFillMode", fill.picture_fill_mode.as_ref()),
    ] {
        if let Some(value) = value {
            paint.insert(key.to_string(), Value::String(value.clone()));
        }
    }
    // resolved picture-fill source: pass through only parser-minted embedded
    // schemes (data:/blob:) so a hand-crafted input cannot smuggle an external
    // URL to the canvas image resolver
    if let Some(src) = fill
        .picture_src
        .as_ref()
        .filter(|src| src.starts_with("data:") || src.starts_with("blob:"))
    {
        paint.insert("pictureSrc".to_string(), Value::String(src.clone()));
    }
    for (key, value) in [
        ("pictureSrcRect", fill.picture_src_rect.as_ref()),
        ("pictureTile", fill.picture_tile.as_ref()),
        ("pictureStretchRect", fill.picture_stretch_rect.as_ref()),
    ] {
        if let Some(value) = value {
            paint.insert(key.to_string(), value.clone());
        }
    }
    if let Some(opacity) = fill.picture_opacity.and_then(Number::from_f64) {
        paint.insert("pictureOpacity".to_string(), Value::Number(opacity));
    }
    if let Some(index) = fill.theme_ref_index {
        paint.insert("themeRefIndex".to_string(), Value::Number(index.into()));
    }
    (!paint.is_empty()).then_some(Value::Object(paint))
}

fn shape_stroke_paint(stroke: Option<&ShapeStrokeIn>) -> Option<Value> {
    let stroke = stroke?;
    let mut paint = serde_json::Map::new();
    for (key, value) in [
        ("color", stroke.color.as_ref()),
        ("dash", stroke.dash.as_ref()),
        ("compound", stroke.compound.as_ref()),
        ("alignment", stroke.alignment.as_ref()),
        ("cap", stroke.cap.as_ref()),
        ("join", stroke.join.as_ref()),
    ] {
        if let Some(value) = value {
            paint.insert(key.to_string(), Value::String(value.clone()));
        }
    }
    for (key, value) in [("width", stroke.width), ("miterLimit", stroke.miter_limit)] {
        if let Some(value) = value.and_then(Number::from_f64) {
            paint.insert(key.to_string(), Value::Number(value));
        }
    }
    if !stroke.custom_dash.is_empty() {
        paint.insert(
            "customDash".to_string(),
            Value::Array(
                stroke
                    .custom_dash
                    .iter()
                    .filter_map(|value| Number::from_f64(*value).map(Value::Number))
                    .collect(),
            ),
        );
    }
    if let Some(value) = &stroke.head_end {
        paint.insert("headEnd".to_string(), value.clone());
    }
    if let Some(value) = &stroke.tail_end {
        paint.insert("tailEnd".to_string(), value.clone());
    }
    (!paint.is_empty()).then_some(Value::Object(paint))
}

fn shape_stroke(stroke: Option<&ShapeStrokeIn>) -> Option<ShapeStrokePrimitive> {
    let stroke = stroke?;
    let width = stroke.width.unwrap_or(1.0);
    if width <= 0.0 {
        return None;
    }
    Some(ShapeStrokePrimitive {
        color: stroke
            .color
            .clone()
            .unwrap_or_else(|| "#000000".to_string()),
        width: px(width),
        dash: stroke
            .dash
            .clone()
            .filter(|d| !d.is_empty() && d != "solid"),
    })
}

fn shape_transform(transform: Option<&ShapeTransformIn>) -> Option<ShapeTransformPrimitive> {
    let transform = transform?;
    let rotation = transform.rotation.filter(|r| *r != 0.0).map(px);
    let flip_h = transform.flip_h == Some(true);
    let flip_v = transform.flip_v == Some(true);
    if rotation.is_none() && !flip_h && !flip_v {
        return None;
    }
    Some(ShapeTransformPrimitive {
        rotation,
        flip_h,
        flip_v,
    })
}

fn scale_shape_path(path: &[ShapePathCommand], frag: &ShapeFragmentIn) -> Vec<ShapePathCommand> {
    path.iter()
        .map(|cmd| match cmd {
            ShapePathCommand::Move { x, y } => ShapePathCommand::Move {
                x: scale_shape_x(x, frag),
                y: scale_shape_y(y, frag),
            },
            ShapePathCommand::Line { x, y } => ShapePathCommand::Line {
                x: scale_shape_x(x, frag),
                y: scale_shape_y(y, frag),
            },
            ShapePathCommand::Quad { cpx, cpy, x, y } => ShapePathCommand::Quad {
                cpx: scale_shape_x(cpx, frag),
                cpy: scale_shape_y(cpy, frag),
                x: scale_shape_x(x, frag),
                y: scale_shape_y(y, frag),
            },
            ShapePathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => ShapePathCommand::Cubic {
                cp1x: scale_shape_x(cp1x, frag),
                cp1y: scale_shape_y(cp1y, frag),
                cp2x: scale_shape_x(cp2x, frag),
                cp2y: scale_shape_y(cp2y, frag),
                x: scale_shape_x(x, frag),
                y: scale_shape_y(y, frag),
            },
            ShapePathCommand::Close => ShapePathCommand::Close,
        })
        .collect()
}

fn scale_shape_x(n: &Number, frag: &ShapeFragmentIn) -> Number {
    px(frag.x + num_f64(n) * frag.width)
}

fn scale_shape_y(n: &Number, frag: &ShapeFragmentIn) -> Number {
    px(frag.y + num_f64(n) * frag.height)
}

fn plot_chart_from(chart: &ChartIn) -> PlotChart<'_> {
    PlotChart {
        chart_type: &chart.chart_type,
        title: chart.title.as_deref(),
        legend: chart.legend.as_ref().map(|legend| PlotLegend {
            position: legend.position.as_deref(),
            visible: legend.visible,
        }),
        value_axis: chart
            .axes
            .as_ref()
            .and_then(|axes| axes.value.as_ref())
            .map(|axis| PlotAxisRange {
                min: axis.min,
                max: axis.max,
            }),
        axis_titles: PlotAxisTitles::default(),
        series: chart.series.iter().map(plot_series_from).collect(),
        plot_groups: chart
            .plot_groups
            .iter()
            .map(|group| PlotGroup {
                chart_type: group.chart_type.as_deref(),
                grouping: group.grouping.as_deref(),
                series: group
                    .series
                    .iter()
                    .map(|series| {
                        let mut plotted = plot_series_from(series);
                        plotted.labels = plot_labels_from(
                            group.data_labels.as_ref(),
                            series.data_labels.as_ref(),
                        );
                        for point in &mut plotted.points {
                            point.labels = plotted.labels;
                        }
                        plotted
                    })
                    .collect(),
                overlap: group.overlap,
                gap_width: group.gap_width,
                hole_size: group.hole_size,
                first_slice_angle: group.first_slice_angle,
                vary_colors: group.vary_colors,
                scatter_style: group.scatter_style.as_deref(),
                radar_style: group.radar_style.as_deref(),
                bubble_scale: group.bubble_scale,
                size_represents: group.size_represents.as_deref(),
                wireframe: group.wireframe,
                hi_low_lines: group.hi_low_lines,
                up_down_bars: group.up_down_bars,
                markers: group.marker,
                axis_ids: group.axis_ids.iter().map(String::as_str).collect(),
            })
            .collect(),
        axes: chart.axis_list.iter().map(plot_axis_from).collect(),
        text: PlotChartText {
            chart: plot_text_from(chart.text.as_ref()),
            title: plot_text_from(chart.title_text.as_ref()),
            legend: plot_text_from(
                chart
                    .legend
                    .as_ref()
                    .and_then(|legend| legend.text.as_ref()),
            ),
        },
    }
}

fn plot_text_from(text: Option<&ChartTextIn>) -> PlotTextStyle<'_> {
    text.map(|text| PlotTextStyle {
        font: text.font.as_deref(),
        size_pt: text.size_pt,
        bold: text.bold,
        italic: text.italic,
        color: text.color.as_deref(),
    })
    .unwrap_or_default()
}

/// Merges a series `dataLabels` over its plot group's, matching the shared
/// geometry's inheritance.
fn plot_labels_from<'a>(
    group: Option<&'a ChartDataLabelsIn>,
    series: Option<&'a ChartDataLabelsIn>,
) -> Option<PlotDataLabels<'a>> {
    let levels = [series, group];
    levels.iter().copied().flatten().next()?;
    if levels
        .iter()
        .copied()
        .flatten()
        .find_map(|labels| labels.delete)
        == Some(true)
    {
        return None;
    }
    let switch = |read: fn(&ChartDataLabelsIn) -> Option<bool>| {
        series
            .and_then(read)
            .or_else(|| group.and_then(read))
            .unwrap_or(false)
    };
    let text = |read: fn(&ChartDataLabelsIn) -> Option<&str>| {
        series.and_then(read).or_else(|| group.and_then(read))
    };
    Some(PlotDataLabels {
        show_value: switch(|labels| labels.show_value),
        show_category_name: switch(|labels| labels.show_category_name),
        show_series_name: switch(|labels| labels.show_series_name),
        show_percent: switch(|labels| labels.show_percent),
        show_legend_key: switch(|labels| labels.show_legend_key),
        show_bubble_size: switch(|labels| labels.show_bubble_size),
        separator: text(|labels| labels.separator.as_deref()),
        position: text(|labels| labels.position.as_deref()),
        number_format: text(|labels| labels.number_format.as_deref()),
        text: plot_text_from(
            series
                .and_then(|labels| labels.text.as_ref())
                .or_else(|| group.and_then(|labels| labels.text.as_ref())),
        ),
    })
}

fn plot_axis_from(axis: &ChartAxisIn) -> PlotAxis<'_> {
    PlotAxis {
        id: axis.id.as_deref(),
        kind: axis
            .axis_type
            .as_deref()
            .map(PlotAxisKind::from_name)
            .unwrap_or(PlotAxisKind::Value),
        range: PlotAxisRange {
            min: axis.min,
            max: axis.max,
        },
        log_base: axis.logarithmic_base,
        reversed: axis.reversed,
        major_unit: axis.major_unit,
        minor_unit: axis.minor_unit,
        major_tick_mark: axis.major_tick_mark.as_deref(),
        minor_tick_mark: axis.minor_tick_mark.as_deref(),
        major_gridlines: axis.major_gridlines,
        minor_gridlines: axis.minor_gridlines,
        number_format: axis.number_format.as_deref(),
        position: axis.position.as_deref(),
        title: axis.title.as_deref(),
        hidden: axis.hidden,
        text: plot_text_from(axis.text.as_ref()),
    }
}

fn plot_series_from(series: &ChartSeriesIn) -> PlotSeries<'_> {
    PlotSeries {
        name: series.name.as_deref(),
        categories: &series.categories,
        values: &series.values,
        color: series.color.as_deref(),
        points: series
            .points
            .iter()
            .map(|point| PlotPoint {
                index: point.index,
                value: point.value,
                color: point.color.as_deref(),
                marker: plot_marker_from(point.marker.as_ref()),
                label: point.label.as_deref(),
                explosion: point.explosion,
                labels: None,
            })
            .collect(),
        grouping: series.grouping.as_deref(),
        marker: plot_marker_from(series.marker.as_ref()),
        x_values: &series.x_values,
        bubble_sizes: &series.bubble_sizes,
        smooth: series.smooth,
        labels: plot_labels_from(None, series.data_labels.as_ref()),
    }
}

fn plot_marker_from(marker: Option<&Value>) -> Option<PlotMarker> {
    marker.map(|marker| PlotMarker {
        size: marker.get("size").and_then(Value::as_f64),
        symbol: marker
            .get("symbol")
            .and_then(Value::as_str)
            .and_then(PlotMarkerSymbol::from_name),
    })
}

fn emit_chart_fragment(prims: &mut Vec<Primitive>, frag: &ChartFragmentIn, block: &ChartBlockIn) {
    let stamp_from = prims.len();
    let block_ref = BlockRef::of(&frag.block_id);
    let chart = plot_chart_from(&block.chart);
    let mut attrs = block_ref.attrs();
    attrs.doc_start = frag
        .doc_start
        .or(frag.pm_start)
        .or(block.doc_start)
        .or(block.pm_start);
    attrs.doc_end = frag
        .doc_end
        .or(frag.pm_end)
        .or(block.doc_end)
        .or(block.pm_end);
    attrs.sdt = sdt_attrs_from_groups(&block.sdt_groups);
    attrs.sdt_path = sdt_path_from_groups(&block.sdt_groups);
    attrs.chart = Some(ChartA11yAttrs {
        label: chart_aria_label(&chart),
    });
    attrs.aria_label = block.chart.title.clone();
    attrs.aria_description = block.chart.description.clone();
    attrs.decorative = block.chart.decorative.filter(|decorative| *decorative);

    let width = if frag.width > 0.0 {
        frag.width
    } else {
        block.width
    };
    let height = if frag.height > 0.0 {
        frag.height
    } else {
        block.height
    };

    plot_chart_into(
        &chart,
        PlotRect {
            x: frag.x,
            y: frag.y,
            w: width,
            h: height,
        },
        &mut PrimitiveSink {
            prims,
            attrs: &attrs,
        },
    );
    stamp_sdt_range(&mut prims[stamp_from..], &block.sdt_groups, false);
}

/// Translates plot ops into display-list primitives as they are emitted.
struct PrimitiveSink<'a> {
    prims: &'a mut Vec<Primitive>,
    attrs: &'a DocAttrs,
}

impl PlotSink for PrimitiveSink<'_> {
    fn push_op(&mut self, op: PlotOp) -> bool {
        let (prims, attrs) = (&mut *self.prims, self.attrs);
        match op {
            PlotOp::Rect { x, y, w, h, fill } => prims.push(Primitive::Rect(RectPrimitive {
                x: px(x),
                y: px(y),
                w: px(w),
                h: px(h),
                fill,
                attrs: attrs.clone(),
            })),
            PlotOp::Text {
                text,
                x,
                baseline_y,
                width,
                font,
                color,
            } => prims.push(Primitive::Text(TextRunPrimitive {
                text,
                x: px(x),
                baseline_y: px(baseline_y),
                width: px(width),
                paint_clip: None,
                font: font.css(),
                color,
                letter_spacing: None,
                word_spacing: None,
                rtl: None,
                opacity: None,
                rotation_deg: None,
                horizontal_scale: None,
                all_caps: false,
                small_caps: false,
                hidden: false,
                text_shadow: None,
                text_outline: false,
                emphasis_mark: None,
                text_effect: None,
                attrs: attrs.clone(),
            })),
            PlotOp::Line {
                x1,
                y1,
                x2,
                y2,
                color,
                width,
            } => prims.push(Primitive::Line(LinePrimitive {
                x1: px(x1),
                y1: px(y1),
                x2: px(x2),
                y2: px(y2),
                stroke_width: px(width),
                color,
                dash: None,
                role: None,
                ..LinePrimitive::contract_defaults()
            })),
            PlotOp::Path {
                x,
                y,
                w,
                h,
                commands,
                fill,
                stroke,
            } => prims.push(Primitive::Shape(ShapePrimitive {
                x: px(x),
                y: px(y),
                w: px(w),
                h: px(h),
                geometry_path: commands.into_iter().map(shape_path_command).collect(),
                fill: Some(fill),
                stroke: stroke.map(|stroke| ShapeStrokePrimitive {
                    color: stroke.color,
                    width: px(stroke.width),
                    dash: None,
                }),
                transform: None,
                decorative: false,
                attrs: attrs.clone(),
            })),
        }
        true
    }
}

fn shape_path_command(command: GeometryPathCommand) -> ShapePathCommand {
    match command {
        GeometryPathCommand::Move { x, y } => ShapePathCommand::Move { x: px(x), y: px(y) },
        GeometryPathCommand::Line { x, y } => ShapePathCommand::Line { x: px(x), y: px(y) },
        GeometryPathCommand::Quad { cpx, cpy, x, y } => ShapePathCommand::Quad {
            cpx: px(cpx),
            cpy: px(cpy),
            x: px(x),
            y: px(y),
        },
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => ShapePathCommand::Cubic {
            cp1x: px(cp1x),
            cp1y: px(cp1y),
            cp2x: px(cp2x),
            cp2y: px(cp2y),
            x: px(x),
            y: px(y),
        },
        GeometryPathCommand::Close => ShapePathCommand::Close,
    }
}

/// Paints one text-box fragment: the container's fill and border edges at the
/// box's page rect, then the inner paragraphs at the content origin. The box
/// behaves like a border-box, so border and padding sit inside the fragment
/// rect and content starts at `x + borderWidth + margin`.
fn emit_text_box_fragment(
    prims: &mut Vec<Primitive>,
    frag: &TextBoxFragmentIn,
    block: &TextBoxBlockIn,
    measure: &TextBoxExtentIn,
    ctx: &RenderCtx<'_>,
) {
    let stamp_from = prims.len();
    let block_ref = BlockRef::of(&frag.block_id);

    // Container fill: a fragment-sized rect behind the content, carrying the
    // text box's own doc range.
    if let Some(fill) = &block.fill_color {
        let mut fill_attrs = block_ref.attrs();
        fill_attrs.doc_start = frag.pm_start;
        fill_attrs.doc_end = frag.pm_end;
        prims.push(Primitive::Rect(RectPrimitive {
            x: px(frag.x),
            y: px(frag.y),
            w: px(frag.width),
            h: px(frag.height),
            fill: fill.clone(),
            attrs: fill_attrs,
        }));
    }

    // border: four edges of the box, inset by half the stroke so a
    // box-sizing:border-box border sits inside the rect (CSS draws it inward)
    let border_w = block.outline_width.unwrap_or(0.0);
    if border_w > 0.0 {
        let color = block
            .outline_color
            .clone()
            .unwrap_or_else(|| "#000000".to_string());
        let h = border_w / 2.0;
        let l = frag.x + h;
        let t = frag.y + h;
        let r = frag.x + frag.width - h;
        let b = frag.y + frag.height - h;
        let mut edge = |x1: f64, y1: f64, x2: f64, y2: f64| {
            let style = display_border_style(block.outline_style.as_deref());
            prims.push(Primitive::Line(LinePrimitive {
                x1: px(x1),
                y1: px(y1),
                x2: px(x2),
                y2: px(y2),
                stroke_width: px(border_w),
                color: color.clone(),
                dash: border_dash(style, border_w),
                role: Some(LineRole::Border),
                border_style: Some(style),
                border_owner: Some(BorderOwner::TextBox),
                ..LinePrimitive::contract_defaults()
            }));
        };
        edge(l, t, r, t); // top
        edge(r, t, r, b); // right
        edge(l, b, r, b); // bottom
        edge(l, t, l, b); // left
    }

    // The inner width excludes margins but not the border.
    let margins = block.margins.unwrap_or(DEFAULT_TEXTBOX_MARGINS);
    let content_x = frag.x + border_w + margins.left;
    let content_top = frag.y + border_w + margins.top;
    let inner_width = (frag.width - margins.left - margins.right).max(0.0);

    let mut y_offset = 0.0;
    for (i, para) in block.content.iter().enumerate() {
        let Some(pm) = measure.inner_measures.get(i) else {
            continue;
        };
        let para_y = content_top + y_offset;
        let synthetic = ParagraphFragmentIn {
            block_id: para.id.clone(),
            x: content_x,
            y: para_y,
            width: inner_width,
            height: pm.total_height,
            from_line: 0,
            to_line: pm.lines.len(),
            pm_start: para.pm_start,
            pm_end: para.pm_end,
            carried_from_prev: None,
            carried_to_next: None,
        };
        emit_paragraph_fragment(
            prims, &synthetic, para, pm, ctx, content_x, para_y, None, None, true, true,
        );
        y_offset += pm.total_height;
    }
    stamp_sdt_range(&mut prims[stamp_from..], &block.sdt_groups, false);
}

// ---------------------------------------------------------------------------
// tables
// ---------------------------------------------------------------------------

/// A cell resolved onto the column grid.
struct GridCell {
    row_index: usize,
    cell_index: usize,
    column_index: usize,
    col_span: usize,
    row_span: usize,
    x: f64,
    width: f64,
}

/// Resolves every cell onto the column grid and gives it a pixel `x` and width
/// from `column_widths`.
fn compute_cell_grid(block: &TableBlockIn, column_widths: &[f64]) -> Vec<GridCell> {
    // rtl tables (`w:bidiVisual`) mirror x so logical column 0 lands rightmost
    let bidi = block.bidi == Some(true);
    let table_width: f64 = column_widths.iter().sum();

    let mut occupied: HashMap<usize, Vec<usize>> = HashMap::new();
    let mut out = Vec::new();
    for (row_index, row) in block.rows.iter().enumerate() {
        let occ = occupied.remove(&row_index).unwrap_or_default();
        let is_occ = |c: usize| occ.contains(&c);
        let mut column_index = 0;
        while is_occ(column_index) {
            column_index += 1;
        }
        for (cell_index, cell) in row.cells.iter().enumerate() {
            let col_span = cell.col_span.unwrap_or(1).max(1) as usize;
            let row_span = cell.row_span.unwrap_or(1).max(1) as usize;
            let mut x = 0.0;
            for c in 0..column_index {
                x += column_widths.get(c).copied().unwrap_or(0.0);
            }
            let mut width = 0.0;
            for c in 0..col_span {
                width += column_widths.get(column_index + c).copied().unwrap_or(0.0);
            }
            if bidi {
                x = table_width - x - width;
            }
            out.push(GridCell {
                row_index,
                cell_index,
                column_index,
                col_span,
                row_span,
                x,
                width,
            });
            if row_span > 1 {
                for r in row_index + 1..row_index + row_span {
                    occupied
                        .entry(r)
                        .or_default()
                        .extend(column_index..column_index + col_span);
                }
            }
            column_index += col_span;
            while is_occ(column_index) {
                column_index += 1;
            }
        }
    }
    out
}

/// Cumulative row offsets, each rounded to a whole pixel so shared borders land
/// on the same device line. Length is `rows + 1`.
fn row_y_positions(rows: &[TableRowExtentIn]) -> Vec<f64> {
    let mut out = Vec::with_capacity(rows.len() + 1);
    let mut y: f64 = 0.0;
    for r in rows {
        out.push(y.round());
        y += r.height;
    }
    out.push(y.round());
    out
}

pub(crate) fn table_total_width(measure: &TableExtentIn) -> f64 {
    if measure.total_width > 0.0 {
        measure.total_width
    } else {
        measure.column_widths.iter().sum()
    }
}

fn nested_table_x_offset(block: &TableBlockIn, measure: &TableExtentIn, content_width: f64) -> f64 {
    let table_width = table_total_width(measure);
    match block.justification.as_deref() {
        Some("center") => ((content_width - table_width) / 2.0).max(0.0),
        Some("right") => (content_width - table_width).max(0.0),
        _ => block.indent.unwrap_or(0.0).max(0.0),
    }
}

fn clip_number(value: &Option<Number>) -> f64 {
    value.as_ref().map(num_f64).unwrap_or(0.0)
}

fn apply_clip_group(attrs: &mut DocAttrs, id: String, rect: ClipRect) {
    let clip = if let Some(existing) = attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.clip.as_ref())
    {
        let left = clip_number(&existing.x).max(clip_number(&rect.x));
        let top = clip_number(&existing.y).max(clip_number(&rect.y));
        let right = (clip_number(&existing.x) + clip_number(&existing.w))
            .min(clip_number(&rect.x) + clip_number(&rect.w));
        let bottom = (clip_number(&existing.y) + clip_number(&existing.h))
            .min(clip_number(&rect.y) + clip_number(&rect.h));
        ClipRect {
            x: Some(px(left)),
            y: Some(px(top)),
            w: Some(px((right - left).max(0.0))),
            h: Some(px((bottom - top).max(0.0))),
        }
    } else {
        rect
    };
    attrs.clip_group = Some(ClipGroupMetadata {
        id: Some(id),
        clip: Some(clip),
        opacity: None,
    });
}

fn table_metadata(
    table_id: &str,
    frag: &TableFragmentIn,
    block: &TableBlockIn,
    measure: &TableExtentIn,
    header_row_count: usize,
) -> TableMetadata {
    TableMetadata {
        table_id: table_id.to_string(),
        row_start: frag.row_start as u64,
        row_end: frag.row_end as u64,
        row_count: block.rows.len() as u64,
        column_count: measure.column_widths.len() as u64,
        header_row_count: (header_row_count > 0).then_some(header_row_count as u64),
        caption: block.caption.clone(),
        description: block.description.clone(),
        parent_table_id: None,
    }
}

/// Paints the row window this page shows of a table: the repeated header band,
/// the visible rows, re-emitted vertical merges, collapsed shared borders, and
/// the cut edges that close the fragment where a page break sliced it.
pub(crate) fn emit_table_fragment(
    prims: &mut Vec<Primitive>,
    frag: &TableFragmentIn,
    block: &TableBlockIn,
    measure: &TableExtentIn,
    ctx: &RenderCtx<'_>,
) {
    let stamp_from = prims.len();
    let table_id = block_key(&frag.block_id);
    let row_tops = row_y_positions(&measure.rows);
    let grid = compute_cell_grid(block, &measure.column_widths);

    let carried = frag.carried_from_prev == Some(true);
    let header_row_count = if carried {
        frag.header_row_count.unwrap_or(0)
    } else {
        0
    };
    let authored_header_count = block
        .rows
        .iter()
        .take_while(|row| row.is_header == Some(true))
        .count();
    let semantic_header_count = authored_header_count.max(header_row_count);
    let mut header_height = 0.0;
    for r in 0..header_row_count.min(measure.rows.len()) {
        header_height += measure.rows[r].height;
    }

    let win_top =
        row_tops.get(frag.row_start).copied().unwrap_or(0.0) + frag.clip_top.unwrap_or(0.0);
    let to_frag_y = |full_y: f64| header_height + (full_y - win_top);
    let visible_height = if frag.clip_bottom.is_some() {
        frag.height.round()
    } else {
        to_frag_y(row_tops.get(frag.row_end).copied().unwrap_or(0.0))
    };

    // Clip band in page coordinates; every emitted rect and text clips to it,
    // so a row sliced by the page break cannot paint past the fragment.
    let clip_top_y = frag.y;
    let clip_bottom_y = frag.y + visible_height;

    let block_ref = BlockRef::of(&frag.block_id);
    let table_revision = whole_table_revision(block);
    if let Some(rev) = table_revision.clone() {
        let mut attrs = block_ref.attrs();
        attrs.structural_revision = Some(rev.clone());
        prims.push(Primitive::Rect(RectPrimitive {
            x: px(frag.x + STRUCTURAL_CHANGE_BAR_OFFSET_X),
            y: px(frag.y),
            w: px(STRUCTURAL_CHANGE_BAR_WIDTH),
            h: px(visible_height),
            fill: structural_color(rev.kind).to_string(),
            attrs,
        }));
    }

    // rows visible in this fragment: repeated header rows first (their own
    // coordinate space above the windowed body), then the body window
    struct VisibleRow {
        row_index: usize,
        frag_y: f64,
        is_first_in_fragment: bool,
    }
    let mut visible: Vec<VisibleRow> = Vec::new();
    if header_row_count > 0 {
        let mut hy = 0.0;
        for r in 0..header_row_count.min(measure.rows.len()) {
            visible.push(VisibleRow {
                row_index: r,
                frag_y: hy,
                is_first_in_fragment: r == 0,
            });
            hy += measure.rows[r].height;
        }
    }
    for row_index in frag.row_start..frag.row_end.min(block.rows.len()) {
        let is_first_in_fragment = if header_row_count > 0 {
            false
        } else {
            carried && row_index == frag.row_start && frag.clip_top.unwrap_or(0.0) == 0.0
        };
        visible.push(VisibleRow {
            row_index,
            frag_y: to_frag_y(row_tops.get(row_index).copied().unwrap_or(0.0)),
            is_first_in_fragment,
        });
    }

    if table_revision.is_none() {
        for vr in &visible {
            let Some(row) = block.rows.get(vr.row_index) else {
                continue;
            };
            let Some(rev) = row_structural_revision(row, vr.row_index) else {
                continue;
            };
            let full_top = frag.y + vr.frag_y;
            let row_h = row_tops.get(vr.row_index + 1).copied().unwrap_or(0.0)
                - row_tops.get(vr.row_index).copied().unwrap_or(0.0);
            let t = full_top.max(clip_top_y);
            let b = (full_top + row_h).min(clip_bottom_y);
            if b - t <= 0.0 {
                continue;
            }
            let mut attrs = block_ref.attrs();
            attrs.structural_revision = Some(rev.clone());
            prims.push(Primitive::Rect(RectPrimitive {
                x: px(frag.x + STRUCTURAL_CHANGE_BAR_OFFSET_X),
                y: px(t),
                w: px(STRUCTURAL_CHANGE_BAR_WIDTH),
                h: px(b - t),
                fill: structural_color(rev.kind).to_string(),
                attrs,
            }));
        }
    }

    // vertically-merged cells whose restart row is on an earlier fragment but
    // whose span reaches into this one re-paint clipped (not selectable — they
    // carry no doc positions, matching data-vmerge-continuation)
    struct CellPaint<'a> {
        g: &'a GridCell,
        cell_y: f64,
        cell_h: f64,
        is_first_row: bool,
        selectable: bool,
    }
    let mut paints: Vec<CellPaint> = Vec::new();

    for g in &grid {
        if g.row_span <= 1 || g.row_index >= frag.row_start {
            continue;
        }
        if g.row_index + g.row_span <= frag.row_start {
            continue;
        }
        if header_row_count > 0 && g.row_index < header_row_count {
            continue; // already drawn by the header pass
        }
        let mut span_height = 0.0;
        for r in g.row_index..(g.row_index + g.row_span).min(row_tops.len() - 1) {
            span_height += row_tops[r + 1] - row_tops[r];
        }
        paints.push(CellPaint {
            g,
            cell_y: to_frag_y(row_tops.get(g.row_index).copied().unwrap_or(0.0)),
            cell_h: span_height,
            is_first_row: false,
            selectable: false,
        });
    }

    for vr in &visible {
        let row_h = row_tops.get(vr.row_index + 1).copied().unwrap_or(0.0)
            - row_tops.get(vr.row_index).copied().unwrap_or(0.0);
        for g in grid.iter().filter(|g| g.row_index == vr.row_index) {
            let mut cell_h = row_h;
            if g.row_span > 1 {
                cell_h = 0.0;
                for r in g.row_index..(g.row_index + g.row_span).min(row_tops.len() - 1) {
                    cell_h += row_tops[r + 1] - row_tops[r];
                }
            }
            paints.push(CellPaint {
                g,
                cell_y: vr.frag_y,
                cell_h,
                is_first_row: g.row_index == 0 || vr.is_first_in_fragment,
                selectable: true,
            });
        }
    }

    let col_count = measure.column_widths.len();
    let bidi = block.bidi == Some(true);

    // per cell: background, collapsed borders, then content — clipped to the
    // fragment window
    for p in &paints {
        let cell_stamp_from = prims.len();
        let cell = &block.rows[p.g.row_index].cells[p.g.cell_index];
        let cx = frag.x + p.g.x;
        let cy = frag.y + p.cell_y;
        // The outer left border insets cell content by its width.
        let is_first_col = if bidi {
            p.g.column_index + p.g.col_span >= col_count
        } else {
            p.g.column_index == 0
        };
        // Grid position stamped on every primitive inside this cell. A vertical
        // merge continuation keeps the anchor cell's row and column and flags
        // itself so consumers know it is a re-paint.
        let is_header = p.g.row_index < semantic_header_count
            || block.rows[p.g.row_index].is_header == Some(true);
        let cell_id = format!("{table_id}-r{}-c{}", p.g.row_index, p.g.column_index);
        let header_ids = if is_header {
            Vec::new()
        } else {
            (0..semantic_header_count)
                .filter_map(|header_row| {
                    grid.iter()
                        .find(|cell| {
                            cell.row_index == header_row
                                && p.g.column_index >= cell.column_index
                                && p.g.column_index < cell.column_index + cell.col_span
                        })
                        .map(|cell| {
                            format!("{table_id}-r{}-c{}", cell.row_index, cell.column_index)
                        })
                })
                .collect()
        };
        let cell_ref = TableCellRef {
            row: p.g.row_index as u64,
            col: p.g.column_index as u64,
            row_span: p.g.row_span as u64,
            col_span: p.g.col_span as u64,
            continuation: if p.selectable { None } else { Some(true) },
            cell_id: Some(cell_id),
            is_header: is_header.then_some(true),
            repeated_header: (carried && p.g.row_index < header_row_count).then_some(true),
            no_wrap: cell.no_wrap.filter(|value| *value),
            header_ids,
            owns_top_border: (p.is_first_row
                && cell
                    .borders
                    .as_ref()
                    .and_then(|borders| borders.top.as_ref())
                    .is_some())
            .then_some(true),
            owns_right_border: cell
                .borders
                .as_ref()
                .and_then(|borders| borders.right.as_ref())
                .is_some()
                .then_some(true),
            owns_bottom_border: cell
                .borders
                .as_ref()
                .and_then(|borders| borders.bottom.as_ref())
                .is_some()
                .then_some(true),
            owns_left_border: (is_first_col
                && cell
                    .borders
                    .as_ref()
                    .and_then(|borders| borders.left.as_ref())
                    .is_some())
            .then_some(true),
        };
        let clip = |top: f64, bottom: f64| -> Option<(f64, f64)> {
            let t = top.max(clip_top_y);
            let b = bottom.min(clip_bottom_y);
            if b - t <= 0.0 { None } else { Some((t, b)) }
        };

        if let Some(bg) = &cell.background
            && let Some((t, b)) = clip(cy, cy + p.cell_h)
        {
            let mut bg_attrs = block_ref.attrs();
            bg_attrs.cell = Some(cell_ref.clone());
            prims.push(Primitive::Rect(RectPrimitive {
                x: px(cx),
                y: px(t),
                w: px(p.g.width),
                h: px(b - t),
                fill: bg.clone(),
                attrs: bg_attrs,
            }));
        }

        if let Some(marker) = &cell.tracked_marker {
            let parent_row_revision_id = row_parent_revision_id(&block.rows[p.g.row_index]);
            if marker.info.revision_id != parent_row_revision_id
                && let Some((t, b)) = clip(cy, cy + CELL_STRUCTURAL_BAR_HEIGHT)
            {
                let rev = structural_revision(
                    &marker.info,
                    StructuralRevisionScope::Cell,
                    marker.kind,
                    Some(p.g.row_index as u64),
                    Some(p.g.column_index as u64),
                );
                let mut attrs = block_ref.attrs();
                attrs.cell = Some(cell_ref.clone());
                attrs.structural_revision = Some(rev.clone());
                prims.push(Primitive::Rect(RectPrimitive {
                    x: px(cx),
                    y: px(t),
                    w: px(p.g.width),
                    h: px(b - t),
                    fill: structural_color(rev.kind).to_string(),
                    attrs,
                }));
            }
        }

        if let Some(borders) = &cell.borders {
            // shared-edge collapse: every cell owns bottom and right; top and
            // left draw only on the table's outer boundary
            let mut push_edge = |x1: f64, y1: f64, x2: f64, y2: f64, e: &BorderEdgeIn| {
                if !border_visible(e) {
                    return;
                }
                // clip vertically to the window band
                let (y1c, y2c) = (y1.max(clip_top_y), y2.min(clip_bottom_y));
                if y1 != y2 && y2c - y1c <= 0.0 {
                    return;
                }
                if y1 == y2 && (y1 < clip_top_y || y1 > clip_bottom_y) {
                    return;
                }
                let style = display_border_style(e.style.as_deref());
                let width = e.width.unwrap_or(1.0);
                // explicit ownership: the owning grid cell rides on the line so
                // consumers associate borders exactly (no geometric fallback)
                let line_attrs = DocAttrs {
                    cell: Some(cell_ref.clone()),
                    ..DocAttrs::default()
                };
                prims.push(Primitive::Line(LinePrimitive {
                    x1: px(x1),
                    y1: px(if y1 == y2 { y1 } else { y1c }),
                    x2: px(x2),
                    y2: px(if y1 == y2 { y2 } else { y2c }),
                    stroke_width: px(width),
                    color: e.color.clone().unwrap_or_else(|| "#000000".to_string()),
                    dash: border_dash(style, width),
                    role: Some(LineRole::TableBorder),
                    border_style: Some(style),
                    border_owner: Some(BorderOwner::Cell),
                    attrs: line_attrs,
                    ..LinePrimitive::contract_defaults()
                }));
            };
            if p.is_first_row
                && let Some(e) = &borders.top
            {
                push_edge(cx, cy, cx + p.g.width, cy, e);
            }
            if let Some(e) = &borders.right {
                push_edge(cx + p.g.width, cy, cx + p.g.width, cy + p.cell_h, e);
            }
            if let Some(e) = &borders.bottom {
                push_edge(cx, cy + p.cell_h, cx + p.g.width, cy + p.cell_h, e);
            }
            if is_first_col && let Some(e) = &borders.left {
                push_edge(cx, cy, cx, cy + p.cell_h, e);
            }
        }

        emit_cell_content(
            prims,
            cell,
            measure,
            &CellPaintRef::from(p.g),
            cx,
            cy,
            p.cell_h,
            p.is_first_row,
            is_first_col,
            clip_top_y,
            clip_bottom_y,
            ctx,
            p.selectable,
            &cell_ref,
            &block_ref,
        );

        let cell_clip_top = cy.max(clip_top_y);
        let cell_clip_bottom = (cy + p.cell_h).min(clip_bottom_y);
        let cell_clip = ClipRect {
            x: Some(px(cx)),
            y: Some(px(cell_clip_top)),
            w: Some(px(p.g.width)),
            h: Some(px((cell_clip_bottom - cell_clip_top).max(0.0))),
        };
        let clip_id = format!("clip-{table_id}-r{}-c{}", p.g.row_index, p.g.column_index);
        for primitive in &mut prims[cell_stamp_from..] {
            if let Some(attrs) = doc_attrs_mut(primitive) {
                apply_clip_group(attrs, clip_id.clone(), cell_clip.clone());
            }
        }
    }

    // cut edges close a fragment at a page break, one rule per column so
    // per-column styles / colSpans / borderless columns are respected;
    // `only_spanning` limits a clean row boundary to cells crossing the edge
    let mut draw_cut_edge = |cut_row: usize, bottom: bool, top_y: f64, only_spanning: bool| {
        for g in &grid {
            if g.row_index > cut_row || g.row_index + g.row_span - 1 < cut_row {
                continue;
            }
            if only_spanning {
                let crosses = if bottom {
                    g.row_index + g.row_span - 1 > cut_row
                } else {
                    g.row_index < cut_row
                };
                if !crosses {
                    continue;
                }
            }
            let cell = &block.rows[g.row_index].cells[g.cell_index];
            let spec = cell.borders.as_ref().and_then(|b| {
                if bottom {
                    b.bottom.as_ref()
                } else {
                    b.top.as_ref()
                }
            });
            let Some(spec) = spec else { continue };
            if !border_visible(spec) {
                continue;
            }
            let bw = spec.width.unwrap_or(1.0);
            let style = display_border_style(spec.style.as_deref());
            // the bottom edge draws upward, sitting just inside the cut
            let y = frag.y + if bottom { top_y - bw } else { top_y };
            // ownership metadata: the cut rule closes this grid cell's column
            // band at the fragment edge (borderOwner stays Fragment)
            let line_attrs = DocAttrs {
                cell: Some(TableCellRef {
                    row: g.row_index as u64,
                    col: g.column_index as u64,
                    row_span: g.row_span as u64,
                    col_span: g.col_span as u64,
                    continuation: None,
                    cell_id: Some(format!("{table_id}-r{}-c{}", g.row_index, g.column_index)),
                    is_header: None,
                    repeated_header: None,
                    no_wrap: None,
                    header_ids: Vec::new(),
                    owns_top_border: None,
                    owns_right_border: None,
                    owns_bottom_border: None,
                    owns_left_border: None,
                }),
                ..DocAttrs::default()
            };
            prims.push(Primitive::Line(LinePrimitive {
                x1: px(frag.x + g.x),
                y1: px(y),
                x2: px(frag.x + g.x + g.width),
                y2: px(y),
                stroke_width: px(bw),
                color: spec.color.clone().unwrap_or_else(|| "#000000".to_string()),
                dash: border_dash(style, bw),
                role: Some(LineRole::TableCut),
                border_style: Some(style),
                border_owner: Some(BorderOwner::Fragment),
                attrs: line_attrs,
                ..LinePrimitive::contract_defaults()
            }));
        }
    };
    if frag.clip_top.unwrap_or(0.0) > 0.0 {
        draw_cut_edge(frag.row_start, false, header_height, false);
    } else if carried {
        draw_cut_edge(frag.row_start, false, header_height, true);
    }
    if frag.clip_bottom.is_some() {
        draw_cut_edge(frag.row_end.saturating_sub(1), true, visible_height, false);
    } else if frag.carried_to_next == Some(true) {
        draw_cut_edge(frag.row_end.saturating_sub(1), true, visible_height, true);
    }
    let metadata = table_metadata(&table_id, frag, block, measure, semantic_header_count);
    for primitive in &mut prims[stamp_from..] {
        // table border/cut lines carry ownership metadata too — doc_attrs_mut
        // deliberately excludes lines from the generic stamping passes, so the
        // fragment-identity stamp reaches them through this explicit arm
        let attrs = match primitive {
            Primitive::Line(line) => match line.role {
                Some(LineRole::TableBorder | LineRole::TableCut) => Some(&mut line.attrs),
                _ => None,
            },
            other => doc_attrs_mut(other),
        };
        if let Some(attrs) = attrs {
            if let Some(inner) = &mut attrs.table {
                if inner.table_id != table_id && inner.parent_table_id.is_none() {
                    inner.parent_table_id = Some(table_id.clone());
                }
            } else {
                attrs.table = Some(metadata.clone());
            }
        }
    }
    stamp_sdt_range(&mut prims[stamp_from..], &block.sdt_groups, false);
}

/// Stacks a cell's paragraphs and nested tables with Word's spacing collapse.
///
/// Adjacent paragraphs collapse `spacing.after` against the next
/// `spacing.before`, a nested table flows after the previous paragraph's
/// after-spacing, and a trailing after-spacing acts as the content box's bottom
/// padding. Drawn border widths and padding inset the box on the sides this
/// cell actually paints, and the resulting box is what `w:vAlign` measures its
/// slack against.
#[allow(clippy::too_many_arguments)]
fn emit_cell_content(
    prims: &mut Vec<Primitive>,
    cell: &TableCellIn,
    measure: &TableExtentIn,
    p: &CellPaintRef,
    cx: f64,
    cy: f64,
    cell_h: f64,
    is_first_row: bool,
    is_first_col: bool,
    clip_top_y: f64,
    clip_bottom_y: f64,
    ctx: &RenderCtx<'_>,
    selectable: bool,
    cell_ref: &TableCellRef,
    block_ref: &BlockRef,
) {
    let Some(row_measure) = measure.rows.get(p.row_index) else {
        return;
    };
    let Some(cell_measure) = row_measure.cells.get(p.cell_index) else {
        return;
    };
    let pad_left = cell.padding.and_then(|pd| pd.left).unwrap_or(7.0);
    let pad_top = cell.padding.and_then(|pd| pd.top).unwrap_or(1.0);
    let pad_right = cell.padding.and_then(|pd| pd.right).unwrap_or(7.0);
    let pad_bottom = cell.padding.and_then(|pd| pd.bottom).unwrap_or(1.0);
    let content_width = (p.width - pad_left - pad_right).max(0.0);

    // Drawn border widths inset the cell content box.
    let edge_w = |e: &Option<BorderEdgeIn>| -> f64 {
        e.as_ref()
            .filter(|b| border_visible(b))
            .map(|b| b.width.unwrap_or(1.0))
            .unwrap_or(0.0)
    };
    let (border_left, border_top, border_bottom) = match &cell.borders {
        Some(b) => (
            if is_first_col { edge_w(&b.left) } else { 0.0 },
            if is_first_row { edge_w(&b.top) } else { 0.0 },
            edge_w(&b.bottom),
        ),
        None => (0.0, 0.0, 0.0),
    };

    // Paragraph spacing collapses within the cell flow.
    let mut block_tops: Vec<f64> = Vec::with_capacity(cell.blocks.len());
    let mut stack_cursor = 0.0_f64;
    let mut prev_after = 0.0_f64;
    for (i, blk) in cell.blocks.iter().enumerate() {
        match (blk, cell_measure.blocks.get(i)) {
            (BlockIn::Paragraph(pb), Some(MeasureIn::Paragraph(pm))) => {
                let spacing = pb.attrs.as_ref().and_then(|a| a.spacing);
                let before = spacing.and_then(|s| s.before).unwrap_or(0.0);
                let after = spacing.and_then(|s| s.after).unwrap_or(0.0);
                stack_cursor += prev_after.max(before);
                block_tops.push(stack_cursor);
                stack_cursor += pm
                    .lines
                    .iter()
                    .map(|l| l.line_height + l.float_skip_before.unwrap_or(0.0))
                    .sum::<f64>();
                prev_after = after;
            }
            (BlockIn::Table(_), Some(MeasureIn::Table(tm))) => {
                stack_cursor += prev_after;
                block_tops.push(stack_cursor);
                stack_cursor += tm.total_height;
                prev_after = 0.0;
            }
            (BlockIn::Image(_), Some(MeasureIn::Image(image))) => {
                stack_cursor += prev_after;
                block_tops.push(stack_cursor);
                stack_cursor += image.height;
                prev_after = 0.0;
            }
            (BlockIn::TextBox(_), Some(MeasureIn::TextBox(text_box))) => {
                stack_cursor += prev_after;
                block_tops.push(stack_cursor);
                stack_cursor += text_box.height;
                prev_after = 0.0;
            }
            (BlockIn::Shape(_), Some(MeasureIn::Shape(sm)))
            | (BlockIn::Chart(_), Some(MeasureIn::Chart(sm))) => {
                stack_cursor += prev_after;
                block_tops.push(stack_cursor);
                stack_cursor += sm.height;
                prev_after = 0.0;
            }
            _ => block_tops.push(stack_cursor),
        }
    }
    // a trailing spacing.after becomes the content box's padding-bottom
    let content_height = stack_cursor + prev_after;

    // Content that fills or overflows the cell remains top-aligned.
    let avail = (cell_h - border_top - border_bottom - pad_top - pad_bottom).max(0.0);
    let content_fills = cell_measure.height >= cell_h - 0.5;
    let v_offset = if content_fills {
        0.0
    } else {
        match cell.vertical_align.as_deref() {
            Some("center") => ((avail - content_height) / 2.0).max(0.0),
            Some("bottom") => (avail - content_height).max(0.0),
            _ => 0.0,
        }
    };

    let content_x = cx + border_left + pad_left;
    let content_top = cy + border_top + pad_top + v_offset;

    // Behind-document floats paint below cell content.
    emit_cell_floating_images(
        prims,
        cell,
        cell_measure,
        block_ref,
        content_x,
        content_top,
        content_width,
        clip_top_y,
        clip_bottom_y,
        cell_ref,
        selectable,
        true,
    );

    for (i, cell_block) in cell.blocks.iter().enumerate() {
        let Some(m) = cell_measure.blocks.get(i) else {
            continue;
        };
        if let (BlockIn::Paragraph(pb), MeasureIn::Paragraph(pm)) = (cell_block, m) {
            // cell paragraphs never split; fabricate a whole-paragraph fragment
            let total_height: f64 = pm
                .lines
                .iter()
                .map(|l| l.line_height + l.float_skip_before.unwrap_or(0.0))
                .sum();
            // y of this paragraph's first line = the collapsed-spacing stack
            // offset computed above (block_tops is index-aligned with cell.blocks)
            let para_y = content_top + block_tops[i];
            let synthetic = ParagraphFragmentIn {
                block_id: pb.id.clone(),
                x: content_x,
                y: para_y,
                width: content_width,
                height: total_height,
                from_line: 0,
                to_line: pm.lines.len(),
                pm_start: if selectable { pb.pm_start } else { None },
                pm_end: if selectable { pb.pm_end } else { None },
                carried_from_prev: None,
                carried_to_next: None,
            };
            let before = prims.len();
            emit_paragraph_fragment(
                // cell paragraphs render as ARIA cells in the mirror, not
                // paragraph fragments — no line-range node to stamp
                prims, &synthetic, pb, pm, ctx, content_x, para_y, None, None, true, false,
            );
            postprocess_cell_primitives(
                prims,
                before,
                clip_top_y,
                clip_bottom_y,
                selectable,
                Some(cell_ref),
            );
        } else if let (BlockIn::Table(tb), MeasureIn::Table(tm)) = (cell_block, m) {
            let table_y = content_top + block_tops[i];
            let synthetic = TableFragmentIn {
                block_id: tb.id.clone(),
                x: content_x + nested_table_x_offset(tb, tm, content_width),
                y: table_y,
                width: table_total_width(tm),
                height: tm.total_height,
                row_start: 0,
                row_end: tb.rows.len(),
                clip_top: None,
                clip_bottom: None,
                header_row_count: None,
                carried_from_prev: None,
                carried_to_next: None,
            };
            let before = prims.len();
            emit_table_fragment(prims, &synthetic, tb, tm, ctx);
            // Preserve the nested table's own inner `cell` refs so the mirror can
            // surface its table semantics; only clip to the outer cell fragment
            // and strip doc positions on a vmerge continuation repaint.
            postprocess_cell_primitives(prims, before, clip_top_y, clip_bottom_y, selectable, None);
        } else if let (BlockIn::Image(image), MeasureIn::Image(image_measure)) = (cell_block, m) {
            let image_x = content_x;
            let image_y = content_top + block_tops[i];
            let mut attrs = BlockRef::of(&image.id).attrs();
            if selectable {
                attrs.doc_start = image.pm_start;
                attrs.doc_end = image.pm_end;
            }
            stamp_image_block_attrs(&mut attrs, image, image_x, image_y);
            let rotation = image
                .rotation_deg
                .unwrap_or_else(|| rotation_degrees(image.transform.as_deref()));
            let before = prims.len();
            prims.push(Primitive::Image(ImagePrimitive {
                rel_id: image.src.clone(),
                x: px(image_x),
                y: px(image_y),
                w: px(image_measure.width),
                h: px(image_measure.height),
                rotation_deg: (rotation != 0.0).then(|| px(rotation)),
                opacity: image.opacity.map(px),
                filter: None,
                decorative: image.decorative.unwrap_or(false),
                crop: crop_of_block(image),
                alt_text: capped_alt_text(image.alt.as_deref()),
                attrs,
            }));
            postprocess_cell_primitives(
                prims,
                before,
                clip_top_y,
                clip_bottom_y,
                selectable,
                Some(cell_ref),
            );
        } else if let (BlockIn::TextBox(text_box), MeasureIn::TextBox(text_box_measure)) =
            (cell_block, m)
        {
            let text_box_y = content_top + block_tops[i];
            let synthetic = TextBoxFragmentIn {
                block_id: text_box.id.clone(),
                x: content_x,
                y: text_box_y,
                width: text_box_measure.width.max(0.0),
                height: text_box_measure.height,
                pm_start: selectable.then_some(text_box.pm_start).flatten(),
                pm_end: selectable.then_some(text_box.pm_end).flatten(),
                is_floating: Some(text_box.display_mode.as_deref() == Some("float")),
                z_index: None,
            };
            let before = prims.len();
            emit_text_box_fragment(prims, &synthetic, text_box, text_box_measure, ctx);
            postprocess_cell_primitives(
                prims,
                before,
                clip_top_y,
                clip_bottom_y,
                selectable,
                Some(cell_ref),
            );
        } else if let (BlockIn::Shape(sb), MeasureIn::Shape(sm)) = (cell_block, m) {
            let shape_y = content_top + block_tops[i];
            let synthetic = ShapeFragmentIn {
                block_id: sb.id.clone(),
                x: content_x,
                y: shape_y,
                width: if sm.width > 0.0 { sm.width } else { sb.width },
                height: if sm.height > 0.0 {
                    sm.height
                } else {
                    sb.height
                },
                doc_start: if selectable { sb.doc_start } else { None },
                doc_end: if selectable { sb.doc_end } else { None },
                pm_start: if selectable { sb.pm_start } else { None },
                pm_end: if selectable { sb.pm_end } else { None },
                is_anchored: None,
                z_index: None,
            };
            let before = prims.len();
            emit_shape_fragment(prims, &synthetic, sb, ctx);
            postprocess_cell_primitives(
                prims,
                before,
                clip_top_y,
                clip_bottom_y,
                selectable,
                Some(cell_ref),
            );
        } else if let (BlockIn::Chart(cb), MeasureIn::Chart(cm)) = (cell_block, m) {
            let chart_y = content_top + block_tops[i];
            let synthetic = ChartFragmentIn {
                block_id: cb.id.clone(),
                x: content_x,
                y: chart_y,
                width: if cm.width > 0.0 { cm.width } else { cb.width },
                height: if cm.height > 0.0 {
                    cm.height
                } else {
                    cb.height
                },
                doc_start: if selectable { cb.doc_start } else { None },
                doc_end: if selectable { cb.doc_end } else { None },
                pm_start: if selectable { cb.pm_start } else { None },
                pm_end: if selectable { cb.pm_end } else { None },
            };
            let before = prims.len();
            emit_chart_fragment(prims, &synthetic, cb);
            postprocess_cell_primitives(
                prims,
                before,
                clip_top_y,
                clip_bottom_y,
                selectable,
                Some(cell_ref),
            );
        }
    }

    // Front cell floats paint above cell content.
    emit_cell_floating_images(
        prims,
        cell,
        cell_measure,
        block_ref,
        content_x,
        content_top,
        content_width,
        clip_top_y,
        clip_bottom_y,
        cell_ref,
        selectable,
        false,
    );
}

fn postprocess_cell_primitives(
    prims: &mut Vec<Primitive>,
    start: usize,
    clip_top_y: f64,
    clip_bottom_y: f64,
    selectable: bool,
    cell_ref: Option<&TableCellRef>,
) {
    let mut k = start;
    while k < prims.len() {
        let (top, bottom) = primitive_v_extent(&prims[k]);
        if bottom < clip_top_y || top > clip_bottom_y {
            prims.remove(k);
        } else {
            if !selectable {
                strip_doc_positions(&mut prims[k]);
            }
            if let Some(cell_ref) = cell_ref {
                set_cell_ref(&mut prims[k], cell_ref);
            }
            k += 1;
        }
    }
}

/// Resolves a cell-anchored floating image to a cell-content-relative origin.
///
/// `paragraph_y` is the anchoring paragraph's top within the cell content box;
/// the caller offsets the result into page space. Unlike the page float path
/// there are no `relativeFrom` bands here — the cell content box is the only
/// frame — and the image is clamped inside it.
fn resolve_cell_float_position(
    imr: &ImageRunIn,
    paragraph_y: f64,
    content_width: f64,
) -> (f64, f64) {
    let width = image_layout_width(imr);

    // horizontal: align keyword, else posOffset, else cssFloat=right, else 0
    let mut x = 0.0;
    if let Some(h) = imr.position.as_ref().and_then(|p| p.horizontal.as_ref()) {
        match h.align.as_deref() {
            Some("right") => x = content_width - width,
            Some("left") => x = 0.0,
            Some("center") => x = (content_width - width) / 2.0,
            _ => {
                if let Some(off) = h.pos_offset {
                    x = emu_to_px(off);
                }
            }
        }
    } else if imr.css_float.as_deref() == Some("right") {
        x = content_width - width;
    }

    // vertical: posOffset from the paragraph top, align=top pins to the cell top,
    // otherwise the anchoring paragraph's top
    let mut y = paragraph_y;
    if let Some(v) = imr.position.as_ref().and_then(|p| p.vertical.as_ref()) {
        if let Some(off) = v.pos_offset {
            y = paragraph_y + emu_to_px(off);
        } else if v.align.as_deref() == Some("top") {
            y = 0.0;
        }
    }

    // clamp inside the content box: Math.max(0, Math.min(x, contentWidth - width))
    let x = x.min(content_width - width).max(0.0);
    (x, y)
}

/// Emits a cell's floating image runs, `want_behind` choosing the layer painted
/// under the cell content or the one over it.
///
/// The anchoring `paragraph_y` accumulates the cell's measured block heights
/// with no spacing collapse. Each image clips to the fragment window, carries
/// the cell's grid reference, and keeps its document positions only on a
/// selectable slice — a vertical-merge continuation is a re-paint and carries
/// none.
#[allow(clippy::too_many_arguments)]
fn emit_cell_floating_images(
    prims: &mut Vec<Primitive>,
    cell: &TableCellIn,
    cell_measure: &TableCellExtentIn,
    block_ref: &BlockRef,
    content_x: f64,
    content_top: f64,
    content_width: f64,
    clip_top_y: f64,
    clip_bottom_y: f64,
    cell_ref: &TableCellRef,
    selectable: bool,
    want_behind: bool,
) {
    let mut paragraph_y = 0.0_f64;
    for (i, blk) in cell.blocks.iter().enumerate() {
        let BlockIn::Paragraph(pb) = blk else {
            // Nested tables advance the anchor cursor by their measured height.
            match cell_measure.blocks.get(i) {
                Some(MeasureIn::Table(tm)) => paragraph_y += tm.total_height,
                Some(MeasureIn::Shape(sm)) | Some(MeasureIn::Chart(sm)) => paragraph_y += sm.height,
                _ => {}
            }
            continue;
        };
        for run in &pb.runs {
            let RunIn::Image(imr) = run else { continue };
            if !is_floating_image_run(imr) {
                continue;
            }
            let is_behind = imr.wrap_type.as_deref() == Some("behind");
            if is_behind != want_behind {
                continue;
            }
            let (x, y) = resolve_cell_float_position(imr, paragraph_y, content_width);
            let page_x = content_x + x;
            let page_y = content_top + y;
            let rot = imr
                .rotation_deg
                .unwrap_or_else(|| rotation_degrees(imr.transform.as_deref()));
            let layout_width = image_layout_width(imr);
            let layout_height = image_layout_height(imr);
            let mut attrs = block_ref.attrs();
            attrs.cell = Some(cell_ref.clone());
            // A continuation repaint carries no document positions.
            if selectable {
                attrs.doc_start = imr.pm_start;
                attrs.doc_end = imr.pm_end;
            }
            stamp_image_run_attrs(&mut attrs, imr, page_x, page_y);
            let prim = Primitive::Image(ImagePrimitive {
                rel_id: imr.src.clone(),
                x: px(page_x),
                y: px(page_y),
                w: px(layout_width),
                h: px(layout_height),
                rotation_deg: if rot != 0.0 { Some(px(rot)) } else { None },
                opacity: imr.opacity.map(px),
                filter: None,
                decorative: imr.decorative.unwrap_or(false),
                crop: crop_of(imr),
                alt_text: capped_alt_text(imr.alt.as_deref()),
                attrs,
            });
            // Clip to the fragment window.
            let (top, bottom) = primitive_v_extent(&prim);
            if bottom < clip_top_y || top > clip_bottom_y {
                continue;
            }
            prims.push(prim);
        }
        if let Some(MeasureIn::Paragraph(pm)) = cell_measure.blocks.get(i) {
            paragraph_y += pm.total_height;
        }
    }
}

/// narrow view of CellPaint the content pass needs (avoids borrowing GridCell)
struct CellPaintRef {
    row_index: usize,
    cell_index: usize,
    width: f64,
}

impl<'a> From<&'a GridCell> for CellPaintRef {
    fn from(g: &'a GridCell) -> Self {
        CellPaintRef {
            row_index: g.row_index,
            cell_index: g.cell_index,
            width: g.width,
        }
    }
}

fn primitive_v_extent(p: &Primitive) -> (f64, f64) {
    match p {
        Primitive::Text(t) => {
            let b = num_f64(&t.baseline_y);
            (b - 16.0, b + 4.0)
        }
        // a GlyphRun replaces a TextRunPrimitive at the same baseline; reuse the
        // Text band (relative to a representative glyph's baseline y) so a table
        // cell's row-break clipping decides identically whether fonts are on or
        // off. Marks sit above the baseline (smaller y), so the max glyph y is
        // the base baseline.
        Primitive::GlyphRun(g) => {
            let b = g
                .glyphs
                .iter()
                .map(|gl| gl.y)
                .fold(f64::NEG_INFINITY, f64::max);
            let b = if b.is_finite() { b } else { 0.0 };
            (b - 16.0, b + 4.0)
        }
        Primitive::Rect(r) => (num_f64(&r.y), num_f64(&r.y) + num_f64(&r.h)),
        Primitive::Line(l) => (
            num_f64(&l.y1).min(num_f64(&l.y2)),
            num_f64(&l.y1).max(num_f64(&l.y2)),
        ),
        Primitive::Image(i) => (num_f64(&i.y), num_f64(&i.y) + num_f64(&i.h)),
        Primitive::Shape(s) => (num_f64(&s.y), num_f64(&s.y) + num_f64(&s.h)),
        Primitive::Decoration(d) => (num_f64(&d.y), num_f64(&d.y) + num_f64(&d.h)),
    }
}

/// The document attributes a primitive carries, if its class has any.
pub(crate) fn doc_attrs(p: &Primitive) -> Option<&DocAttrs> {
    match p {
        Primitive::Text(t) => Some(&t.attrs),
        Primitive::GlyphRun(g) => Some(&g.attrs),
        Primitive::Rect(r) => Some(&r.attrs),
        Primitive::Image(i) => Some(&i.attrs),
        Primitive::Shape(s) => Some(&s.attrs),
        Primitive::Decoration(d) => Some(&d.attrs),
        Primitive::Line(_) => None,
    }
}

fn doc_attrs_mut(p: &mut Primitive) -> Option<&mut DocAttrs> {
    match p {
        Primitive::Text(t) => Some(&mut t.attrs),
        Primitive::GlyphRun(g) => Some(&mut g.attrs),
        Primitive::Rect(r) => Some(&mut r.attrs),
        Primitive::Image(i) => Some(&mut i.attrs),
        Primitive::Shape(s) => Some(&mut s.attrs),
        Primitive::Decoration(d) => Some(&mut d.attrs),
        Primitive::Line(_) => None,
    }
}

/// vmerge continuation slices are re-paints of content owned by another
/// fragment — visible but not selectable, so their doc positions are stripped
fn strip_doc_positions(p: &mut Primitive) {
    if let Some(attrs) = doc_attrs_mut(p) {
        attrs.doc_start = None;
        attrs.doc_end = None;
        // A re-painted slice is not a paragraph-id target either; keeping it
        // would expose a duplicate id for the anchor cell.
        attrs.para_id = None;
    }
}

/// stamp the owning cell's grid position on a cell-content primitive (lines
/// carry no DocAttrs and stay untouched)
fn set_cell_ref(p: &mut Primitive, cell: &TableCellRef) {
    if let Some(attrs) = doc_attrs_mut(p) {
        attrs.cell = Some(cell.clone());
    }
}

// ---------------------------------------------------------------------------
// JSON boundary
// ---------------------------------------------------------------------------

/// Pure JSON boundary: `{ measured, options, layout }` in, `DisplayList` JSON
/// out, with no shaping fonts — so every text run emits a
/// [`TextRunPrimitive`]. Native-testable (no `JsValue`); the fonts-aware entry
/// is [`build_display_list_json_with_fonts`].
pub fn build_display_list_json(input: &str) -> Result<String, String> {
    build_display_list_json_with_fonts(input, &ooxml_text::FontStore::new())
}

/// pure JSON boundary that threads a measurement `FontStore`: when the input
/// carries `fontChains` that resolve against `fonts`, text runs are shaped into
/// [`GlyphRunPrimitive`]s; otherwise (empty store, absent chains, or an
/// unresolved run) the run falls back to `TextRunPrimitive`. This is the entry
/// the wasm wrapper in `lib.rs` drives with the module-global measurement fonts.
pub fn build_display_list_json_with_fonts(
    input: &str,
    fonts: &ooxml_text::FontStore,
) -> Result<String, String> {
    let dl = build_display_list_value_with_fonts(input, fonts)?;
    serde_json::to_string(&dl).map_err(|e| format!("serialize: {e}"))
}

/// Typed counterpart to [`build_display_list_json_with_fonts`].
pub fn build_display_list_value_with_fonts(
    input: &str,
    fonts: &ooxml_text::FontStore,
) -> Result<DisplayList, String> {
    let mut wire: Value = serde_json::from_str(input).map_err(|e| format!("parse: {e}"))?;
    normalize_integral_json_numbers(&mut wire);
    let parsed: BuildInput = serde_json::from_value(wire).map_err(|e| format!("parse: {e}"))?;
    Ok(build_display_list(&parsed, fonts))
}

/// Build from the pagination arena already retained by the editing engine.
/// Only display-specific extras (headers/footers, font chains, comments, and
/// the contract version) cross the wasm boundary; measured blocks, options,
/// and layout are injected from typed resident values.
pub fn build_display_list_value_from_resident_with_fonts(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
) -> Result<DisplayList, String> {
    build_display_list_value_from_resident_with_fonts_observed(
        pagination,
        layout,
        extras,
        fonts,
        &mut || {},
    )
}

pub fn build_display_list_value_from_resident_with_fonts_observed(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
    observe_phase: &mut impl FnMut(),
) -> Result<DisplayList, String> {
    let parsed = resident_build_input(pagination, layout, extras)?;
    observe_phase();
    let list = build_display_list(&parsed, fonts);
    observe_phase();
    Ok(list)
}

/// Build and retain the parsed display-input mirror alongside its first list.
/// Incremental engine frames can then refresh only the pages they rebuild.
pub fn build_resident_display_list_with_fonts_observed(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
    observe_phase: &mut impl FnMut(),
) -> Result<(ResidentDisplayInput, DisplayList), String> {
    let input = resident_build_input(pagination, layout, extras)?;
    observe_phase();
    let list = build_display_list(&input, fonts);
    observe_phase();
    Ok((ResidentDisplayInput { input }, list))
}

fn resident_build_input(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
) -> Result<BuildInput, String> {
    let mut wire: serde_json::Map<String, Value> =
        serde_json::from_str(extras).map_err(|e| format!("parse display extras: {e}"))?;
    wire.insert(
        "measured".to_owned(),
        serde_json::to_value(&pagination.measured)
            .map_err(|e| format!("encode resident measured blocks: {e}"))?,
    );
    wire.insert(
        "options".to_owned(),
        serde_json::to_value(&pagination.options)
            .map_err(|e| format!("encode resident layout options: {e}"))?,
    );
    wire.insert(
        "layout".to_owned(),
        serde_json::to_value(layout).map_err(|e| format!("encode resident layout: {e}"))?,
    );
    let mut wire = Value::Object(wire);
    normalize_integral_json_numbers(&mut wire);
    serde_json::from_value(wire).map_err(|e| format!("parse resident display input: {e}"))
}

/// Rebuild only pages dirtied by incremental pagination, retain the remaining
/// typed display pages, and patch absolute body positions on the converged
/// suffix. The caller gates extras/page-count changes before selecting this
/// path; violations widen to a full display build here as a final safeguard.
pub fn build_display_list_value_from_resident_incremental_with_fonts(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
    previous: &DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &HashMap<String, i64>,
) -> Result<DisplayList, String> {
    let parsed = resident_build_input(pagination, layout, extras)?;
    if previous.pages.len() != parsed.layout.pages.len()
        || rebuilt_page_start > rebuilt_page_end
        || rebuilt_page_end > parsed.layout.pages.len()
    {
        return Ok(build_display_list(&parsed, fonts));
    }

    let selected: HashSet<_> = (rebuilt_page_start..rebuilt_page_end).collect();
    let rebuilt = build_display_list_selected(&parsed, fonts, Some(&selected));
    let mut rebuilt_by_index: HashMap<usize, DisplayPage> = rebuilt
        .pages
        .into_iter()
        .map(|page| (page.page_index as usize, page))
        .collect();
    let mut pages = Vec::with_capacity(previous.pages.len());
    for (page_index, previous_page) in previous.pages.iter().enumerate() {
        if let Some(page) = rebuilt_by_index.remove(&page_index) {
            pages.push(page);
            continue;
        }
        let mut page = previous_page.clone();
        page.page_index = page_index as u64;
        if page_index >= rebuilt_page_end {
            shift_page_body_positions(&mut page, position_deltas);
        }
        pages.push(page);
    }
    Ok(DisplayList {
        contract_version: rebuilt.contract_version,
        pages,
    })
}

/// In-place counterpart to
/// [`build_display_list_value_from_resident_incremental_with_fonts`]. Engine
/// sessions already own the previous display arena, so unchanged pages do not
/// need to be deep-cloned for every keystroke. Dirty pages are replaced after
/// they have been built successfully; the converged suffix receives only its
/// absolute-position adjustment.
pub fn update_display_list_value_from_resident_incremental_with_fonts(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
    previous: &mut DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &HashMap<String, i64>,
) -> Result<bool, String> {
    update_display_list_value_from_resident_incremental_with_fonts_observed(
        pagination,
        layout,
        extras,
        fonts,
        previous,
        rebuilt_page_start,
        rebuilt_page_end,
        position_deltas,
        &mut || {},
    )
}

pub fn update_display_list_value_from_resident_incremental_with_fonts_observed(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    extras: &str,
    fonts: &ooxml_text::FontStore,
    previous: &mut DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &HashMap<String, i64>,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    let parsed = resident_build_input(pagination, layout, extras)?;
    observe_phase();
    if previous.pages.len() != parsed.layout.pages.len()
        || rebuilt_page_start > rebuilt_page_end
        || rebuilt_page_end > parsed.layout.pages.len()
    {
        *previous = build_display_list(&parsed, fonts);
        observe_phase();
        return Ok(false);
    }

    let selected: HashSet<_> = (rebuilt_page_start..rebuilt_page_end).collect();
    let rebuilt = build_display_list_selected(&parsed, fonts, Some(&selected));
    observe_phase();
    previous.contract_version = rebuilt.contract_version;
    for page in rebuilt.pages {
        let page_index = page.page_index as usize;
        previous.pages[page_index] = page;
    }
    for (page_index, page) in previous.pages.iter_mut().enumerate().skip(rebuilt_page_end) {
        page.page_index = page_index as u64;
        shift_page_body_positions(page, position_deltas);
    }
    Ok(true)
}

/// Incremental engine path backed by a retained parsed display input. Only
/// rebuilt layout pages and the measured blocks referenced by those pages
/// cross the typed-layout compatibility adapter on each edit.
pub fn update_resident_display_list_incremental_with_fonts_observed(
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    fonts: &ooxml_text::FontStore,
    resident: &mut ResidentDisplayInput,
    previous: &mut DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &HashMap<String, i64>,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    if previous.pages.len() != layout.pages.len()
        || resident.input.layout.pages.len() != layout.pages.len()
    {
        return Ok(false);
    }
    if rebuilt_page_start > rebuilt_page_end || rebuilt_page_end > layout.pages.len() {
        return Err("resident display incremental page range is invalid".to_owned());
    }

    refresh_resident_display_pages(
        &mut resident.input,
        pagination,
        layout,
        rebuilt_page_start..rebuilt_page_end,
    )?;
    observe_phase();

    let selected: HashSet<_> = (rebuilt_page_start..rebuilt_page_end).collect();
    let rebuilt = build_display_list_selected(&resident.input, fonts, Some(&selected));
    observe_phase();
    previous.contract_version = rebuilt.contract_version;
    for page in rebuilt.pages {
        let page_index = page.page_index as usize;
        previous.pages[page_index] = page;
    }
    for (page_index, page) in previous.pages.iter_mut().enumerate().skip(rebuilt_page_end) {
        page.page_index = page_index as u64;
        shift_page_body_positions(page, position_deltas);
    }
    Ok(true)
}

fn refresh_resident_display_pages(
    input: &mut BuildInput,
    pagination: &crate::types::Input,
    layout: &crate::types::Layout,
    rebuilt_pages: std::ops::Range<usize>,
) -> Result<(), String> {
    let mut selected_blocks = HashSet::new();
    for page_index in rebuilt_pages {
        let page: PageIn =
            convert_resident_value(&layout.pages[page_index], "resident display layout page")?;
        for fragment in &page.fragments {
            if let Some(key) = fragment_block_key(fragment) {
                selected_blocks.insert(key);
            }
        }
        input.layout.pages[page_index] = page;
    }

    let current_indices: HashMap<String, usize> = input
        .measured
        .iter()
        .enumerate()
        .filter_map(|(index, measured)| measured_block_key(measured).map(|key| (key, index)))
        .collect();
    let mut pending_blocks = selected_blocks;
    for measured in &pagination.measured {
        let key = crate_block_key(&measured.block);
        if !pending_blocks.remove(&key) {
            continue;
        }
        let index = current_indices
            .get(&key)
            .copied()
            .ok_or_else(|| format!("resident display measured block {key:?} is missing"))?;
        input.measured[index] =
            convert_resident_value(measured, "resident display measured block")?;
    }
    if let Some(key) = pending_blocks.into_iter().next() {
        return Err(format!(
            "resident pagination measured block {key:?} is missing"
        ));
    }
    Ok(())
}

fn convert_resident_value<T: Serialize, U: DeserializeOwned>(
    input: &T,
    label: &str,
) -> Result<U, String> {
    let mut value =
        serde_json::to_value(input).map_err(|error| format!("encode {label}: {error}"))?;
    normalize_integral_json_numbers(&mut value);
    serde_json::from_value(value).map_err(|error| format!("parse {label}: {error}"))
}

fn crate_block_key(block: &crate::types::LayoutBlock) -> String {
    match block {
        crate::types::LayoutBlock::Paragraph(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::Table(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::Image(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::TextBox(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::Shape(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::Chart(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::SectionBreak(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::PageBreak(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::ColumnBreak(value) => crate_block_id_key(&value.id),
        crate::types::LayoutBlock::Unsupported => "unsupported".to_owned(),
    }
}

fn crate_block_id_key(id: &crate::types::BlockId) -> String {
    match id {
        crate::types::BlockId::Str(value) => value.clone(),
        crate::types::BlockId::Num(value) => value.to_string(),
    }
}

fn measured_block_key(measured: &MeasuredBlockIn) -> Option<String> {
    match &measured.block {
        BlockIn::Paragraph(value) => Some(block_key(&value.id)),
        BlockIn::Table(value) => Some(block_key(&value.id)),
        BlockIn::Image(value) => Some(block_key(&value.id)),
        BlockIn::TextBox(value) => Some(block_key(&value.id)),
        BlockIn::Shape(value) => Some(block_key(&value.id)),
        BlockIn::Chart(value) => Some(block_key(&value.id)),
        BlockIn::Unsupported => None,
    }
}

fn fragment_block_key(fragment: &FragmentIn) -> Option<String> {
    match fragment {
        FragmentIn::Paragraph(value) => Some(block_key(&value.block_id)),
        FragmentIn::Table(value) => Some(block_key(&value.block_id)),
        FragmentIn::Image(value) => Some(block_key(&value.block_id)),
        FragmentIn::TextBox(value) => Some(block_key(&value.block_id)),
        FragmentIn::Shape(value) => Some(block_key(&value.block_id)),
        FragmentIn::Chart(value) => Some(block_key(&value.block_id)),
        FragmentIn::Unsupported => None,
    }
}

fn shift_page_body_positions(page: &mut DisplayPage, deltas: &HashMap<String, i64>) {
    for primitive in &mut page.primitives {
        let attrs = match primitive {
            Primitive::Text(value) => &mut value.attrs,
            Primitive::GlyphRun(value) => &mut value.attrs,
            Primitive::Rect(value) => &mut value.attrs,
            Primitive::Line(value) => &mut value.attrs,
            Primitive::Image(value) => &mut value.attrs,
            Primitive::Shape(value) => &mut value.attrs,
            Primitive::Decoration(value) => &mut value.attrs,
        };
        let key = attrs
            .block_key
            .clone()
            .or_else(|| attrs.block_id.as_ref().map(ToString::to_string));
        let Some(delta) = key.as_ref().and_then(|key| deltas.get(key)).copied() else {
            continue;
        };
        attrs.doc_start = attrs.doc_start.map(|value| value + delta);
        attrs.doc_end = attrs.doc_end.map(|value| value + delta);
        attrs.fragment_doc_start = attrs.fragment_doc_start.map(|value| value + delta);
        attrs.fragment_doc_end = attrs.fragment_doc_end.map(|value| value + delta);
        if let Some(widget) = &mut attrs.inline_sdt_widget {
            widget.pos += delta;
        }
    }
}

/// Rewrites losslessly-integral JSON floats as integers.
///
/// Several established display-input fields deserialize as `i64`, but the same
/// values reach this builder as `f64` when they come from typed Rust state
/// rather than through a host with a single number type. Canonicalizing here
/// makes both paths parse identically.
fn normalize_integral_json_numbers(value: &mut Value) {
    match value {
        Value::Array(values) => {
            for value in values {
                normalize_integral_json_numbers(value);
            }
        }
        Value::Object(fields) => {
            for value in fields.values_mut() {
                normalize_integral_json_numbers(value);
            }
        }
        Value::Number(number) if !number.is_i64() && !number.is_u64() => {
            let Some(float) = number.as_f64() else {
                return;
            };
            if float.is_finite()
                && float.fract() == 0.0
                && float >= i64::MIN as f64
                && float <= i64::MAX as f64
            {
                *number = Number::from(float as i64);
            }
        }
        Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resident_adapter_normalizes_integral_json_numbers() {
        let mut value = serde_json::json!({
            "pmStart": 16.0,
            "fractional": 16.25,
            "nested": [3.0, -4.0],
            "fontChains": { "calibri|0|0": [1.0] }
        });
        normalize_integral_json_numbers(&mut value);
        assert!(value["pmStart"].as_i64().is_some());
        assert_eq!(value["fractional"].as_f64(), Some(16.25));
        assert_eq!(value["nested"][0].as_i64(), Some(3));
        assert_eq!(value["nested"][1].as_i64(), Some(-4));
        let chains: HashMap<String, Vec<u32>> =
            serde_json::from_value(value["fontChains"].clone()).unwrap();
        assert_eq!(chains["calibri|0|0"], vec![1]);
    }
    use serde_json::json;

    #[test]
    fn list_marker_emits_as_a_tracked_first_line_primitive() {
        let input = json!({
            "contractVersion": 1,
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "list-item",
                    "runs": [{ "kind": "text", "text": "item", "pmStart": 1, "pmEnd": 5 }],
                    "attrs": {
                        "listMarker": "1.",
                        "listMarkerFontFamily": "Aptos",
                        "listMarkerFontSize": 12,
                        "listMarkerRevision": "ins",
                        "indent": { "left": 48, "hanging": 24 }
                    },
                    "pmStart": 0,
                    "pmEnd": 5
                },
                "measure": {
                    "kind": "paragraph",
                    "totalHeight": 20,
                    "lines": [{
                        "headRun": 0,
                        "headChar": 0,
                        "tailRun": 0,
                        "tailChar": 4,
                        "width": 28,
                        "ascent": 14,
                        "descent": 4,
                        "lineHeight": 20
                    }]
                }
            }],
            "options": {},
            "layout": {
                "pages": [{
                    "number": 1,
                    "size": { "w": 300, "h": 400 },
                    "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                    "fragments": [{
                        "kind": "paragraph",
                        "blockId": "list-item",
                        "x": 20,
                        "y": 20,
                        "width": 260,
                        "height": 20,
                        "fromLine": 0,
                        "toLine": 1,
                        "pmStart": 0,
                        "pmEnd": 5
                    }]
                }]
            }
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let marker = output["pages"][0]["primitives"]
            .as_array()
            .expect("primitive array")
            .iter()
            .find(|primitive| primitive["listMarker"] == true)
            .expect("synthetic list marker");
        assert_eq!(marker["text"], "1.");
        assert_eq!(marker["listMarkerRevision"], "ins");
        assert_eq!(marker["color"], REVISION_INS_COLOR);
        assert!(marker["font"].as_str().unwrap().contains("Aptos"));
    }

    #[test]
    fn authoritative_clusters_leaders_and_transformed_images_serialize_exact_geometry() {
        let input = json!({
            "contractVersion": 1,
            "measured": [
                {
                    "block": {
                        "kind": "paragraph",
                        "id": "p",
                        "runs": [
                            {
                                "kind": "text",
                                "text": "A😀B",
                                "pmStart": 10,
                                "pmEnd": 14,
                                "highlight": "#ffff00",
                                "emphasisMark": "dot"
                            },
                            {
                                "kind": "tab",
                                "pmStart": 14,
                                "pmEnd": 15,
                                "width": 10,
                                "leaderGlyphs": {
                                    "glyph": "·",
                                    "count": 2,
                                    "advance": 4,
                                    "font": "Calibri",
                                    "fontSize": 11,
                                    "color": "#123456"
                                }
                            }
                        ],
                        "attrs": { "tabs": [{ "pos": 1440, "leader": "middleDot" }] },
                        "pmStart": 10,
                        "pmEnd": 15
                    },
                    "measure": {
                        "kind": "paragraph",
                        "totalHeight": 20,
                        "lines": [{
                            "headRun": 0,
                            "headChar": 0,
                            "tailRun": 1,
                            "tailChar": 1,
                            "width": 35,
                            "ascent": 14,
                            "descent": 4,
                            "lineHeight": 20,
                            "runAdvances": [
                                { "runIndex": 0, "startChar": 0, "endChar": 4, "advance": 25, "logicalOrder": 0 },
                                { "runIndex": 1, "startChar": 0, "endChar": 1, "advance": 10, "logicalOrder": 1000000 }
                            ],
                            "clusterAdvances": [
                                { "runIndex": 0, "startChar": 0, "endChar": 1, "advance": 5, "xOffset": 20, "bidiLevel": 0, "logicalOrder": 0 },
                                { "runIndex": 0, "startChar": 1, "endChar": 3, "advance": 13, "xOffset": 7, "bidiLevel": 1, "logicalOrder": 1 },
                                { "runIndex": 0, "startChar": 3, "endChar": 4, "advance": 7, "xOffset": 0, "bidiLevel": 0, "logicalOrder": 2 }
                            ],
                            "bidiSlices": [
                                { "runIndex": 0, "startChar": 3, "endChar": 4, "advance": 7, "bidiLevel": 0, "visualOrder": 0, "logicalOrder": 2 },
                                { "runIndex": 0, "startChar": 1, "endChar": 3, "advance": 13, "bidiLevel": 1, "visualOrder": 1, "logicalOrder": 1 },
                                { "runIndex": 0, "startChar": 0, "endChar": 1, "advance": 5, "bidiLevel": 0, "visualOrder": 2, "logicalOrder": 0 },
                                { "runIndex": 1, "startChar": 0, "endChar": 1, "advance": 10, "bidiLevel": 0, "visualOrder": 3, "logicalOrder": 1000000 }
                            ]
                        }]
                    }
                },
                {
                    "block": {
                        "kind": "image",
                        "id": "img",
                        "src": "rId9",
                        "width": 40,
                        "height": 20,
                        "opacity": 0.4,
                        "rotationDeg": 30,
                        "flipH": true,
                        "rotationBounds": { "width": 44.641, "height": 37.321, "offsetX": 2.321, "offsetY": 8.66 }
                    },
                    "measure": { "kind": "image", "width": 44.641, "height": 37.321 }
                }
            ],
            "options": {},
            "layout": {
                "pages": [{
                    "number": 1,
                    "size": { "w": 300, "h": 400 },
                    "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                    "fragments": [
                        { "kind": "paragraph", "blockId": "p", "x": 20, "y": 20, "width": 260, "height": 20, "fromLine": 0, "toLine": 1, "pmStart": 10, "pmEnd": 15 },
                        { "kind": "image", "blockId": "img", "x": 20, "y": 50, "width": 44.641, "height": 37.321 }
                    ]
                }]
            }
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        assert_eq!(output["contractVersion"], 1);
        let primitives = output["pages"][0]["primitives"]
            .as_array()
            .expect("primitive array");
        let text: Vec<&Value> = primitives
            .iter()
            .filter(|primitive| primitive["kind"] == "text" && primitive["text"] != "··")
            .collect();
        assert_eq!(
            text.iter()
                .map(|p| p["text"].as_str().unwrap())
                .collect::<Vec<_>>(),
            ["B", "😀", "A"]
        );
        assert_eq!(
            text.iter()
                .map(|p| p["width"].as_f64().unwrap())
                .collect::<Vec<_>>(),
            [7.0, 13.0, 5.0]
        );
        assert_eq!(text[1]["docStart"], 11);
        assert_eq!(text[1]["docEnd"], 13);
        assert_eq!(text[1]["logicalOrder"], 1);
        assert_eq!(text[1]["bidiLevel"], 1);
        let leader = primitives
            .iter()
            .find(|primitive| primitive["leaderGlyphs"].is_object())
            .expect("shaped leader primitive");
        assert_eq!(leader["text"], "··");
        assert_eq!(leader["leaderGlyphs"]["advance"], 4);
        let image = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "image")
            .expect("image primitive");
        assert_eq!(image["opacity"], 0.4);
        assert_eq!(image["flipH"], true);
        assert_eq!(image["w"], 44.641);
        assert_eq!(image["contentFrame"]["w"], 40);
    }

    #[test]
    fn page_regions_emit_columns_notes_and_hf_floating_content() {
        let paragraph = |id: &str, text: &str| {
            json!({
                "block": { "kind": "paragraph", "id": id, "runs": [{ "kind": "text", "text": text, "pmStart": 1, "pmEnd": 1 + text.len() }], "pmStart": 1, "pmEnd": 1 + text.len() },
                "measure": { "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": text.len(), "width": 30, "ascent": 11, "descent": 3, "lineHeight": 16 }] }
            })
        };
        let note = paragraph("note-p", "note");
        let input = json!({
            "measured": [],
            "options": {},
            "headersFooters": {
                "titlePg": false,
                "variants": [{
                    "rId": "rH",
                    "kind": "header",
                    "type": "default",
                    "height": 32,
                    "flowHeight": 32,
                    "measured": [
                        {
                            "block": { "kind": "paragraph", "id": "hf-p", "runs": [{ "kind": "image", "src": "logo", "width": 30, "height": 12, "displayMode": "float", "wrapType": "inFront", "opacity": 0.6, "position": { "horizontal": { "relativeTo": "page", "posOffset": 914400 }, "vertical": { "relativeTo": "page", "posOffset": 457200 } } }] },
                            "measure": { "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1, "width": 30, "ascent": 11, "descent": 3, "lineHeight": 16 }] }
                        },
                        {
                            "block": { "kind": "textBox", "id": "hf-box", "width": 100, "height": 16, "fillColor": "#eeeeee", "content": [{ "kind": "paragraph", "id": "box-p", "runs": [{ "kind": "text", "text": "box", "pmStart": 1, "pmEnd": 4 }] }] },
                            "measure": { "kind": "textBox", "width": 100, "height": 16, "innerMeasures": [{ "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 3, "width": 24, "ascent": 11, "descent": 3, "lineHeight": 16 }] }] }
                        }
                    ]
                }]
            },
            "layout": {
                "pages": [{
                    "number": 1,
                    "size": { "w": 300, "h": 400 },
                    "margins": { "top": 40, "right": 20, "bottom": 40, "left": 20, "header": 20 },
                    "columns": { "count": 2, "gap": 20, "separator": true },
                    "noteAreas": [{ "kind": "footnote", "y": 330, "height": 30, "columns": 1, "notes": [{ "id": 7, "blocks": [note["block"].clone()], "measures": [note["measure"].clone()], "height": 16 }] }],
                    "fragments": []
                }]
            }
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let page = &output["pages"][0];
        assert!(
            page["primitives"]
                .as_array()
                .unwrap()
                .iter()
                .any(|primitive| primitive["role"] == "separator")
        );
        assert_eq!(page["noteAreas"][0]["noteIds"], json!([7]));
        assert!(
            page["noteAreas"][0]["primitives"]
                .as_array()
                .unwrap()
                .iter()
                .any(|primitive| primitive["text"] == "note")
        );
        let header = page["header"]["primitives"]
            .as_array()
            .expect("header primitives");
        assert!(header.iter().any(|primitive| primitive["kind"] == "image"
            && primitive["relId"] == "logo"
            && primitive["opacity"] == 0.6));
        assert!(
            header
                .iter()
                .any(|primitive| primitive["kind"] == "text" && primitive["text"] == "box")
        );
    }

    /// Every primitive a note paints addresses the note's own story, so one
    /// carrying no position of its own must stay unpositioned rather than
    /// inherit the body range of the reference mark that anchors the note —
    /// that range would hand a click inside the note a body position under a
    /// note region. The reference label leading every note is exactly such a
    /// primitive: it is presentation, not story content.
    #[test]
    fn note_primitives_never_inherit_the_body_anchor_range() {
        let input = json!({
            "measured": [],
            "options": {},
            "layout": { "pages": [{
                "number": 1,
                "size": { "w": 300, "h": 400 },
                "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                "noteAreas": [{
                    "kind": "footnote", "y": 330, "height": 40, "columns": 1,
                    "notes": [{
                        "id": 7, "anchorDocStart": 3, "anchorDocEnd": 4, "height": 30,
                        "blocks": [{
                            "kind": "paragraph",
                            "id": "note-p",
                            "runs": [
                                { "kind": "text", "text": "1  " },
                                { "kind": "text", "text": "note", "pmStart": 1, "pmEnd": 5 }
                            ],
                            "pmStart": 1,
                            "pmEnd": 6
                        }],
                        "measures": [{ "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 1, "tailChar": 4, "width": 50, "ascent": 11, "descent": 3, "lineHeight": 16 }] }]
                    }]
                }],
                "fragments": []
            }]}
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let primitives = output["pages"][0]["noteAreas"][0]["primitives"]
            .as_array()
            .expect("note primitives");

        let label = &primitives[0];
        assert_eq!(label["text"], "1  ");
        assert_eq!(label["groupId"], "footnote-7");
        assert!(label["docStart"].is_null() && label["docEnd"].is_null());

        let text = &primitives[1];
        assert_eq!(text["text"], "note");
        assert_eq!(text["groupId"], "footnote-7");
        assert_eq!((&text["docStart"], &text["docEnd"]), (&json!(1), &json!(5)));
    }

    #[test]
    fn table_fragments_emit_clip_border_and_accessibility_contracts() {
        let para = |id: &str, text: &str| {
            json!({
                "kind": "paragraph",
                "id": id,
                "runs": [{ "kind": "text", "text": text, "pmStart": 1, "pmEnd": 2 }],
                "pmStart": 1,
                "pmEnd": 2
            })
        };
        let para_measure = || {
            json!({
                "kind": "paragraph",
                "totalHeight": 16,
                "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1, "width": 8, "ascent": 11, "descent": 3, "lineHeight": 16 }]
            })
        };
        let input = json!({
            "measured": [{
                "block": {
                    "kind": "table",
                    "id": "table-1",
                    "caption": "Example",
                    "description": "Two row table",
                    "rows": [
                        {
                            "isHeader": true,
                            "cells": [{
                                "blocks": [para("head", "H")],
                                "borders": {
                                    "top": { "style": "double", "width": 2, "color": "#111111" },
                                    "right": { "style": "dotted", "width": 1, "color": "#222222" },
                                    "bottom": { "style": "dashDot", "width": 1, "color": "#333333" },
                                    "left": { "style": "single", "width": 1, "color": "#444444" }
                                }
                            }]
                        },
                        {
                            "cells": [{
                                "noWrap": true,
                                "blocks": [
                                    para("data", "D"),
                                    { "kind": "textBox", "id": "cell-box", "width": 80, "height": 16, "fillColor": "#eeeeee", "content": [para("cell-box-p", "T")] }
                                ],
                                "borders": {
                                    "right": { "style": "double", "width": 2, "color": "#555555" },
                                    "bottom": { "style": "dashed", "width": 1, "color": "#666666" },
                                    "left": { "style": "single", "width": 1, "color": "#777777" }
                                }
                            }]
                        }
                    ]
                },
                "measure": {
                    "kind": "table",
                    "columnWidths": [120],
                    "totalWidth": 120,
                    "totalHeight": 64,
                    "rows": [
                        { "height": 24, "cells": [{ "width": 120, "height": 24, "blocks": [para_measure()] }] },
                        { "height": 40, "cells": [{ "width": 120, "height": 40, "blocks": [para_measure(), { "kind": "textBox", "width": 80, "height": 16, "innerMeasures": [para_measure()] }] }] }
                    ]
                }
            }],
            "options": {},
            "layout": {
                "pages": [{
                    "number": 2,
                    "size": { "w": 300, "h": 400 },
                    "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                    "fragments": [{
                        "kind": "table",
                        "blockId": "table-1",
                        "x": 20,
                        "y": 20,
                        "width": 120,
                        "height": 64,
                        "rowStart": 1,
                        "rowEnd": 2,
                        "headerRowCount": 1,
                        "carriedFromPrev": true
                    }]
                }]
            }
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let primitives = output["pages"][0]["primitives"]
            .as_array()
            .expect("primitive array");
        assert!(
            primitives
                .iter()
                .any(|primitive| primitive["kind"] == "line"
                    && primitive["borderStyle"] == "double"
                    && primitive["borderOwner"] == "cell")
        );
        let header = primitives
            .iter()
            .find(|primitive| primitive["text"] == "H")
            .expect("repeated header text");
        assert_eq!(header["cell"]["isHeader"], true);
        assert_eq!(header["cell"]["repeatedHeader"], true);
        let data = primitives
            .iter()
            .find(|primitive| primitive["text"] == "D")
            .expect("data text");
        assert_eq!(data["cell"]["noWrap"], true);
        assert_eq!(data["cell"]["headerIds"], json!(["table-1-r0-c0"]));
        assert_eq!(data["table"]["rowStart"], 1);
        assert_eq!(data["table"]["rowEnd"], 2);
        assert_eq!(data["table"]["caption"], "Example");
        assert!(data["clipGroup"]["clip"].is_object());
        assert!(primitives.iter().any(|primitive| primitive["text"] == "T"));
    }

    #[test]
    fn sdt_links_and_resolved_review_metadata_emit_without_fallback() {
        let input = json!({
            "contractVersion": 1,
            "resolvedCommentIds": [4],
            "commentAuthors": [{ "id": "reviewer-1", "name": "Ada", "paletteIndex": 2, "color": "#123456" }],
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "review",
                    "sdtGroups": [
                        { "id": "outer", "sdtType": "repeatingSection", "alias": "Items" },
                        { "id": "inner", "sdtType": "dropDownList", "tag": "choice", "lock": "sdtLocked" }
                    ],
                    "runs": [
                        { "kind": "text", "text": "R", "pmStart": 1, "pmEnd": 2, "commentIds": [4] },
                        { "kind": "text", "text": "A", "pmStart": 2, "pmEnd": 3, "commentIds": [5], "changeAuthor": "Ada", "changeRevisionId": 9, "isInsertion": true },
                        {
                            "kind": "text",
                            "text": "L",
                            "pmStart": 3,
                            "pmEnd": 4,
                            "hyperlink": { "href": "https://example.test", "tooltip": "Open example", "target": "_blank", "history": true },
                            "inlineSdtWidget": {
                                "kind": "checkbox",
                                "controlKind": "dropDownList",
                                "groupId": "inner",
                                "pos": 3,
                                "controlId": 42,
                                "value": "b",
                                "selectedIndex": 1,
                                "listItems": [{ "displayText": "Alpha", "value": "a" }, { "displayText": "Beta", "value": "b" }],
                                "locked": true
                            }
                        }
                    ],
                    "pmStart": 1,
                    "pmEnd": 4
                },
                "measure": {
                    "kind": "paragraph",
                    "totalHeight": 16,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 2, "tailChar": 1, "width": 24, "ascent": 11, "descent": 3, "lineHeight": 16 }]
                }
            }],
            "options": {},
            "layout": { "pages": [{
                "number": 1,
                "size": { "w": 300, "h": 400 },
                "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                "fragments": [{ "kind": "paragraph", "blockId": "review", "x": 20, "y": 20, "width": 260, "height": 16, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 4 }]
            }]}
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let primitives = output["pages"][0]["primitives"]
            .as_array()
            .expect("primitive array");
        let resolved = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "text" && primitive["text"] == "R")
            .expect("resolved comment text");
        assert_eq!(resolved["comment"]["status"], "resolved");
        assert!(!primitives.iter().any(|primitive| {
            primitive["kind"] == "decoration"
                && primitive["deco"] == "comment-range"
                && primitive["docStart"] == 1
        }));
        let active_wash = primitives
            .iter()
            .find(|primitive| {
                primitive["kind"] == "decoration"
                    && primitive["deco"] == "comment-range"
                    && primitive["docStart"] == 2
            })
            .expect("active reviewer wash");
        assert_eq!(active_wash["comment"]["status"], "active");
        assert_eq!(active_wash["comment"]["authorId"], "reviewer-1");
        assert_eq!(active_wash["color"], "rgba(18, 52, 86, 0.15)");
        let linked = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "text" && primitive["text"] == "L")
            .expect("linked widget text");
        assert_eq!(linked["tooltip"], "Open example");
        assert_eq!(linked["linkTitle"], "Open example");
        assert_eq!(linked["inlineSdtWidget"]["controlKind"], "dropDownList");
        assert_eq!(linked["inlineSdtWidget"]["listItems"][1]["value"], "b");
        assert_eq!(linked["sdt"]["groupId"], "inner");
        assert_eq!(linked["sdtPath"][0]["groupId"], "outer");
        assert_eq!(linked["sdtPath"][1]["groupId"], "inner");
    }

    /// Emits field, note, and comment accessibility metadata.
    #[test]
    fn field_note_and_comment_thread_a11y_metadata_emit() {
        let note = json!({
            "block": { "kind": "paragraph", "id": "note-p", "runs": [{ "kind": "text", "text": "1  note body" }] },
            "measure": { "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 12, "width": 60, "ascent": 11, "descent": 3, "lineHeight": 16 }] }
        });
        let input = json!({
            "contractVersion": 1,
            "commentAuthors": [{ "id": "reviewer-1", "name": "Ada", "paletteIndex": 2 }],
            "commentThreads": [{
                "id": 4,
                "authorId": "reviewer-1",
                "authorName": "Ada Lovelace",
                "date": "2026-07-01T10:00:00Z",
                "text": "Please tighten this sentence.",
                "replies": [
                    { "authorName": "Bob", "date": "2026-07-02T09:00:00Z", "text": "Agreed." },
                    { "authorName": "Ada Lovelace", "text": "Done." }
                ]
            }],
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "body",
                    "runs": [
                        { "kind": "text", "text": "C", "pmStart": 1, "pmEnd": 2, "commentIds": [4] },
                        { "kind": "field", "fieldType": "OTHER", "rawType": "PAGEREF", "instruction": " PAGEREF _Toc42 \\h ", "fallback": "7", "pmStart": 2, "pmEnd": 3 },
                        { "kind": "text", "text": "1", "pmStart": 3, "pmEnd": 4, "superscript": true, "footnoteRefId": 7 }
                    ],
                    "pmStart": 1,
                    "pmEnd": 4
                },
                "measure": {
                    "kind": "paragraph",
                    "totalHeight": 16,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 2, "tailChar": 1, "width": 24, "ascent": 11, "descent": 3, "lineHeight": 16 }]
                }
            }],
            "options": {},
            "layout": { "pages": [{
                "number": 1,
                "size": { "w": 300, "h": 400 },
                "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                "noteAreas": [{
                    "kind": "footnote", "y": 330, "height": 30, "columns": 1,
                    "notes": [{ "id": 7, "displayLabel": "1", "anchorDocStart": 3, "anchorDocEnd": 4, "blocks": [note["block"].clone()], "measures": [note["measure"].clone()], "height": 16 }]
                }],
                "fragments": [{ "kind": "paragraph", "blockId": "body", "x": 20, "y": 20, "width": 260, "height": 16, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 4 }]
            }]}
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let page = &output["pages"][0];
        let primitives = page["primitives"].as_array().expect("primitive array");

        // 1) inert field identity on the field-result primitive
        let field = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "text" && primitive["text"] == "7")
            .expect("field result text");
        assert_eq!(field["field"]["category"], "OTHER");
        assert_eq!(field["field"]["type"], "PAGEREF");
        assert_eq!(field["field"]["instruction"], " PAGEREF _Toc42 \\h ");
        // ordinary text never carries field identity
        let plain = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "text" && primitive["text"] == "C")
            .expect("plain text");
        assert!(plain["field"].is_null());

        // 2) note backlinks: body reference mark + region note metadata
        let ref_mark = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "text" && primitive["text"] == "1")
            .expect("footnote reference mark");
        assert_eq!(ref_mark["noteRef"]["kind"], "footnote");
        assert_eq!(ref_mark["noteRef"]["id"], 7);
        let region_notes = page["noteAreas"][0]["notes"]
            .as_array()
            .expect("note backlink metadata");
        assert_eq!(region_notes[0]["id"], 7);
        assert_eq!(region_notes[0]["anchorDocStart"], 3);
        assert_eq!(region_notes[0]["anchorDocEnd"], 4);
        assert_eq!(region_notes[0]["label"], "1");

        // 3) comment-thread announcement metadata joined by comment id
        assert_eq!(plain["comment"]["status"], "active");
        assert_eq!(plain["comment"]["authorId"], "reviewer-1");
        assert_eq!(plain["comment"]["authorName"], "Ada Lovelace");
        assert_eq!(plain["comment"]["date"], "2026-07-01T10:00:00Z");
        assert_eq!(plain["comment"]["text"], "Please tighten this sentence.");
        assert_eq!(plain["comment"]["replyCount"], 2);
        assert_eq!(plain["comment"]["replies"][0]["authorName"], "Bob");
        assert_eq!(plain["comment"]["replies"][0]["text"], "Agreed.");
        assert_eq!(plain["comment"]["replies"][1]["text"], "Done.");
    }

    #[test]
    fn word_page_decoration_shape_and_combo_chart_contracts_emit() {
        let input = json!({
            "contractVersion": 1,
            "measured": [
                {
                    "block": { "kind": "paragraph", "id": "u", "runs": [{ "kind": "text", "text": "wave", "pmStart": 1, "pmEnd": 5, "underline": { "style": "wave", "color": "#123456" } }] },
                    "measure": { "kind": "paragraph", "totalHeight": 16, "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4, "width": 32, "ascent": 11, "descent": 3, "lineHeight": 16 }] }
                },
                {
                    "block": {
                        "kind": "shape",
                        "id": "rich-shape",
                        "shapeType": "rect",
                        "geometryPath": [{ "type": "move", "x": 0, "y": 0 }, { "type": "line", "x": 1, "y": 1 }, { "type": "close" }],
                        "fill": { "type": "gradient", "gradientAngle": 45, "gradientStops": [{ "position": 0, "color": "#ff0000" }, { "position": 1, "color": "#0000ff" }] },
                        "stroke": { "color": "#222222", "width": 2, "dash": "dashDot", "compound": "double", "headEnd": { "type": "triangle" } },
                        "effects": [{ "kind": "glow", "color": "#00ff00", "blurRadius": 4 }],
                        "effectExtent": { "top": 4, "right": 4, "bottom": 4, "left": 4 },
                        "textBodyProperties": { "anchor": "middle", "margins": { "left": 5, "right": 5, "top": 3, "bottom": 3 } },
                        "scene": { "version": 1, "root": { "kind": "group", "id": "group-7" } },
                        "title": "Process box",
                        "description": "A gradient process shape",
                        "decorative": false,
                        "width": 100,
                        "height": 60
                    },
                    "measure": { "kind": "shape", "width": 100, "height": 60 }
                },
                {
                    "block": {
                        "kind": "chart",
                        "id": "combo",
                        "width": 180,
                        "height": 120,
                        "chart": {
                            "type": "chart",
                            "chartType": "column",
                            "title": "Quarterly",
                            "description": "Revenue and trend",
                            "decorative": true,
                            "series": [],
                            "plotGroups": [
                                { "chartType": "column", "grouping": "clustered", "series": [{ "name": "Revenue", "categories": ["Q1", "Q2"], "values": [5, 9], "points": [{ "index": 1, "value": 10, "color": "#abcdef", "label": "Ten" }] }] },
                                { "chartType": "line", "series": [{ "name": "Trend", "categories": ["Q1", "Q2"], "values": [4, 8], "marker": { "size": 6 } }] }
                            ]
                        }
                    },
                    "measure": { "kind": "chart", "width": 180, "height": 120 }
                }
            ],
            "options": {},
            "layout": { "pages": [{
                "number": 7,
                "pageLabel": "vii",
                "sectionId": "sect-2",
                "sectionIndex": 1,
                "sectionPageIndex": 2,
                "sectionPageNumber": 7,
                "size": { "w": 400, "h": 500 },
                "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
                "fragments": [
                    { "kind": "paragraph", "blockId": "u", "x": 20, "y": 20, "width": 360, "height": 16, "fromLine": 0, "toLine": 1 },
                    { "kind": "shape", "blockId": "rich-shape", "x": 20, "y": 50, "width": 100, "height": 60 },
                    { "kind": "chart", "blockId": "combo", "x": 150, "y": 50, "width": 180, "height": 120 }
                ]
            }]}
        });
        let output: Value = serde_json::from_str(
            &build_display_list_json(&input.to_string()).expect("display list builds"),
        )
        .expect("valid display JSON");
        let page = &output["pages"][0];
        assert_eq!(page["pageLabel"], "vii");
        assert_eq!(page["sectionId"], "sect-2");
        assert_eq!(page["sectionPageNumber"], 7);
        let primitives = page["primitives"].as_array().expect("primitive array");
        assert!(primitives.iter().any(|primitive| {
            primitive["kind"] == "decoration"
                && primitive["deco"] == "underline"
                && primitive["style"] == "wave"
        }));
        let shape = primitives
            .iter()
            .find(|primitive| primitive["kind"] == "shape" && primitive["blockKey"] == "rich-shape")
            .expect("rich shape primitive");
        assert_eq!(shape["fillPaint"]["kind"], "gradient");
        assert_eq!(shape["fillPaint"]["stops"][1]["color"], "#0000ff");
        assert_eq!(shape["strokePaint"]["headEnd"]["type"], "triangle");
        assert_eq!(shape["effects"][0]["kind"], "glow");
        assert_eq!(shape["ariaLabel"], "Process box");
        assert_eq!(shape["groupId"], "group-7");
        let chart = primitives
            .iter()
            .find(|primitive| primitive["chart"]["label"].is_string())
            .expect("combo chart primitive");
        assert!(
            chart["chart"]["label"]
                .as_str()
                .unwrap()
                .contains("combo chart")
        );
        assert_eq!(chart["ariaDescription"], "Revenue and trend");
        assert_eq!(chart["decorative"], true);
        assert!(
            primitives
                .iter()
                .any(|primitive| primitive["fill"] == "#abcdef")
        );
    }

    #[test]
    fn a_group_delete_suppresses_a_series_that_sets_other_switches() {
        let group = ChartDataLabelsIn {
            delete: Some(true),
            ..Default::default()
        };
        let series = ChartDataLabelsIn {
            show_value: Some(true),
            ..Default::default()
        };
        assert!(plot_labels_from(Some(&group), Some(&series)).is_none());

        let restoring = ChartDataLabelsIn {
            delete: Some(false),
            show_value: Some(true),
            ..Default::default()
        };
        let resolved = plot_labels_from(Some(&group), Some(&restoring)).unwrap();
        assert!(resolved.show_value);

        let series_only = ChartDataLabelsIn {
            delete: Some(true),
            ..Default::default()
        };
        assert!(plot_labels_from(None, Some(&series_only)).is_none());
    }
}
