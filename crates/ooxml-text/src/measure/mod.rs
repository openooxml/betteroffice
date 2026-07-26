//! Paragraph measurement: one `paragraph` block and a font set in, a
//! `{ kind: "paragraph", lines, totalHeight }` extent out.
//!
//! [`measure_paragraph`] is pure and panic-free on untrusted input. Anything
//! it will not measure comes back as a [`MeasureError`], never a panic, and
//! [`MeasureError::Unsupported`] stringifies as `"UNSUPPORTED: <reason>"` —
//! the signal for the host to measure that one block with the browser.
//!
//! # Line spans
//!
//! A [`TypesetRowOut`] addresses its text as `headRun..=tailRun` (inclusive
//! run indices) and `headChar..tailChar` — `headChar` inclusive, `tailChar`
//! exclusive. Character positions are **UTF-16 code-unit** offsets into the
//! run's original string, the indexing the wire format uses. Rust text is
//! UTF-8, so every emitted index is a sum of `char::len_utf16` and always
//! lands on a `char` boundary: a split surrogate is unrepresentable.
//! [`TypesetRowSegmentOut`] and the advance metadata index the same way.
//!
//! # What is refused
//!
//! Four categories, all of them rules:
//!
//! - **Resource clamps** on file-derived numbers and counts — font size,
//!   letter spacing, horizontal scale, spacing, indents, image dimensions,
//!   tab-stop / zone / segment / run / line counts, over-long run text.
//!   Degenerate values are refused rather than fed into layout arithmetic.
//! - **Host-contract misses** — no chain, or an empty chain, registered for
//!   a `(family, bold, italic)` key the block uses. Resolving a family to
//!   font bytes is the host's job; measurement never guesses a face.
//! - **Malformed runs** — a mandatory-break control character (`\n`, `\r`,
//!   `\t`, `\v`, `\f`, `U+0085`, `U+2028`, `U+2029`) inside a `text` or
//!   `field` run's string. A well-formed DOCX carries those as `lineBreak`
//!   and `tab` runs, so this fires only on corrupt or hostile input. An
//!   unrecognized `run.kind` is refused the same way.
//! - **Non-paragraph blocks** — `block.kind != "paragraph"`. Tables, block
//!   images, text boxes and breaks are measured by other code paths and are
//!   never routed here.
//!
//! A character no font in the chain covers is *not* refused: it shapes as
//! the chain's terminal font's `.notdef`, which is a real advance width.
//!
//! # Measurement rules
//!
//! - **Metrics come from font tables**: ascent and descent are
//!   `size_px × usWinAscent/upem` and `size_px × usWinDescent/upem`, and the
//!   single-line basis is their sum plus GDI external leading
//!   ([`crate::word_metrics`]).
//! - **Wrapping is greedy** with a half-pixel tolerance: a line takes the
//!   last break opportunity that fits. Opportunities are UAX-14
//!   ([`crate::line_break`]), so consecutive spaces collapse into a single
//!   opportunity and a soft hyphen permits a break.
//! - **An overlong unbreakable word** fills the room left on the current
//!   line and then hard-breaks, at least one shaped cluster per line.
//! - **Trailing whitespace stays** on the line it ends, and its advance is
//!   included in that line's `width` — words are measured with their
//!   trailing space and never trimmed.
//! - **Widths come from shaping whole same-style subranges**, so kerning and
//!   ligature advances straddling a hard-break cut are attributed to the
//!   line before the cut.
//! - **`allCaps` uppercases before shaping**, `smallCaps` shapes lowercase
//!   as small capitals, and `horizontalScale` multiplies glyph advances:
//!   measurement reflects what a painter draws. `letterSpacing` is added
//!   once per gap between shaped clusters within a word — never between
//!   words, and not scaled by `horizontalScale`.
//! - **Line height** follows `spacing.lineRule` (`exact` / `atLeast` /
//!   `auto`) and `lineUnit` (`multiplier` / `px`), defaulting to single
//!   spacing. An empty paragraph floors at 1.15 × the font size under the
//!   `auto` and `atLeast` rules.
//! - **`totalHeight`** sums line heights and float skips, plus
//!   `spacing.before` and `spacing.after`.
//!
//! # Tabs, fields, images, list markers
//!
//! Tab widths resolve against the 720-twip grid overlaid with the
//! paragraph's declared stops; `end` and `center` stops subtract the width
//! of the runs that follow the tab. Field runs measure their `fallback`
//! text (`"1"` when absent or empty) at the run's family, size, bold and
//! italic, with no caps and no letter spacing.
//!
//! An inline image adds its declared width to the line advance and grows the
//! line box by its column-fitted height plus wrap distances: alone on a line
//! it takes a descent buffer above and below, flowing with text it seats on
//! the baseline. A `topAndBottom` or block image takes its own line at its
//! declared height plus wrap distances (default 6px, never column-fitted),
//! adds no width, and opens a fresh line after it. An anchored floating
//! image is positioned by the host, so it contributes neither width nor
//! height — but its declared width still counts toward the following-runs
//! width after a tab.
//!
//! A visible list marker narrows the first line by its footprint, and only
//! when the paragraph's hanging indent is exactly zero.
//!
//! # Float exclusion zones
//!
//! `floatingZones` and `paragraphYOffset` place the paragraph in the float
//! group's coordinate space. Intersecting zones are resolved per line at the
//! running Y with a fixed probe height of `pt_to_px(defaults.fontSize)` —
//! never the line's own fonts, which are unknown until the line closes. That
//! running Y advances by each finalized line's *text* height, so image
//! growth reaches `totalHeight` but not the next line's zone probe.
//!
//! Lines beside a zone shrink by its margins and report `leftOffset` /
//! `rightOffset`. When the room left is under `MIN_WRAP_SEGMENT_WIDTH`, the
//! line hops below the obstruction and the skipped pixels are reported as
//! `floatSkipBefore`. Zones that split a line into strips emit `segments`;
//! a two-way split applies only to a single-text-run line, and any other
//! line emits no segments at all.
//!
//! # Bidirectional text
//!
//! Text is split into UBA level runs ([`crate::bidi`]) before shaping, so
//! every shaped segment is a single directional run with an explicit
//! direction. Everything else stays in logical order: break opportunities,
//! `headChar`/`tailChar` spans, and widths — a line's width is the
//! direction-independent sum of its segment advances. `w:bidi` on the
//! paragraph and `w:rtl` on a run only force the UBA base direction, which
//! changes how neutral characters segment, never the sums.

mod floats;
mod input;
mod line_filler;
mod list_marker;
mod prepare;
mod tabs;

pub use input::{
    AttrsIn, BlockIn, CompatIn, DefaultsIn, FloatSegmentIn, FloatZoneIn, IndentIn, MeasureInput,
    RunIn, SpacingIn, TabStopIn,
};

use crate::font_store::{FontId, FontStore};

/// Why measurement refused an input.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MeasureError {
    /// Outside what this engine measures: a resource clamp, a missing font
    /// chain, a malformed run, or a non-paragraph block. Displays as
    /// `"UNSUPPORTED: <reason>"`; the host measures the block with the
    /// browser instead.
    Unsupported(String),
    /// The request itself is unusable — unparseable JSON, a font id absent
    /// from the store, a shaping failure.
    Invalid(String),
}

impl std::fmt::Display for MeasureError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MeasureError::Unsupported(reason) => write!(f, "UNSUPPORTED: {reason}"),
            MeasureError::Invalid(reason) => write!(f, "invalid: {reason}"),
        }
    }
}

impl std::error::Error for MeasureError {}

/// One typeset line. `headRun`/`tailRun` are inclusive run indices;
/// `headChar` is inclusive and `tailChar` exclusive, both UTF-16 code-unit
/// offsets (see module docs). Unset optional fields are omitted from the
/// JSON rather than serialized as null or zero.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRowOut {
    pub head_run: u32,
    pub head_char: u32,
    pub tail_run: u32,
    pub tail_char: u32,
    pub width: f32,
    pub ascent: f32,
    pub descent: f32,
    pub line_height: f32,
    /// Px from the content left edge (floats); `Some` only when > 0.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_offset: Option<f32>,
    /// Px from the content right edge (floats); `Some` only when > 0.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_offset: Option<f32>,
    /// Split strips for centered floating exclusions.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub segments: Option<Vec<TypesetRowSegmentOut>>,
    /// Vertical px inserted before this line to skip past obstructing
    /// floats; painters render it as top margin, `totalHeight` includes it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub float_skip_before: Option<f32>,
    /// Exact advances for run slices, emitted in visual paint order.
    /// `Some` only under `authoritativeShaping`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_advances: Option<Vec<TypesetRunAdvanceOut>>,
    /// Exact shaped cluster advances and visual x offsets.
    /// `Some` only under `authoritativeShaping`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cluster_advances: Option<Vec<TypesetClusterAdvanceOut>>,
    /// Bidi slices keep logical identity separate from visual paint order.
    /// `Some` only under `authoritativeShaping`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi_slices: Option<Vec<TypesetBidiSliceOut>>,
}

/// Advance of one run's slice of a line. A run split by direction yields one
/// entry per visual piece; `logical_order` recovers logical sequence.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRunAdvanceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub logical_order: u32,
}

/// One shaped cluster: its advance and its visual x from the line start.
/// Clusters are indivisible, so a cluster may span several characters.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetClusterAdvanceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub x_offset: f32,
    pub bidi_level: u8,
    pub logical_order: u32,
}

/// Every line contribution — text clusters and atomic runs alike — with both
/// orderings, so a painter can lay out visually without losing logical
/// identity. Odd `bidi_level` is RTL.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetBidiSliceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub bidi_level: u8,
    pub visual_order: u32,
    pub logical_order: u32,
}

/// One strip of a line split by a segmenting float zone. Spans index exactly
/// as [`TypesetRowOut`]; offsets and widths are pixels.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRowSegmentOut {
    pub head_run: u32,
    pub head_char: u32,
    pub tail_run: u32,
    pub tail_char: u32,
    pub left_offset: f32,
    pub available_width: f32,
    pub width: f32,
}

/// Measured paragraph envelope.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphExtentOut {
    /// Always `"paragraph"`.
    pub kind: &'static str,
    pub lines: Vec<TypesetRowOut>,
    pub total_height: f32,
}

/// Resource limits on file-derived counts. Exceeding any of them is
/// `UNSUPPORTED`, so a hostile document costs a browser fallback rather than
/// unbounded work.
pub(crate) const MAX_RUNS: usize = 10_000;
pub(crate) const MAX_RUN_TEXT_BYTES: usize = 1_000_000;
pub(crate) const MAX_LINES: usize = 100_000;
pub(crate) const MAX_TAB_STOPS: usize = 1_000;
pub(crate) const MAX_FLOAT_ZONES: usize = 200;
pub(crate) const MAX_ZONE_SEGMENTS: usize = 100;

/// Converts points to CSS pixels (1pt = 96/72px).
pub(crate) fn pt_to_px(pt: f32) -> f32 {
    pt * 96.0 / 72.0
}

/// Measure one paragraph block against `input.maxWidth`.
///
/// Pure and panic-free on untrusted input; every validation failure or
/// refused feature comes back as a [`MeasureError`], never a panic.
///
/// A paragraph with no runs, or with a single whitespace-only text run,
/// short-circuits to one zero-width line at the resolved line height, always
/// measured with the regular (`|0|0`) face of the resolved family.
/// `suppressEmptyParagraphHeight` makes that line zero-height instead.
pub fn measure_paragraph(
    store: &FontStore,
    input: &MeasureInput,
) -> Result<ParagraphExtentOut, MeasureError> {
    if input.block.kind != "paragraph" {
        return Err(MeasureError::Unsupported(format!(
            "block kind {:?}",
            input.block.kind
        )));
    }
    if !input.max_width.is_finite() {
        return Err(MeasureError::Unsupported("non-finite maxWidth".to_string()));
    }
    input::validate_pt_size(input.defaults.font_size, "defaults.fontSize")?;

    let runs = &input.block.runs;
    if runs.len() > MAX_RUNS {
        return Err(MeasureError::Unsupported(format!(
            "too many runs ({} > {MAX_RUNS})",
            runs.len()
        )));
    }

    let attrs = input.block.attrs.as_ref();
    let spacing = attrs.and_then(|a| a.spacing.as_ref());
    if let Some(sp) = spacing {
        sp.validate()?;
    }
    if let Some(ind) = attrs.and_then(|a| a.indent.as_ref()) {
        ind.validate()?;
    }
    if let Some(tabs) = attrs.and_then(|a| a.tabs.as_deref()) {
        input::validate_tabs(tabs)?;
    }
    let zones = input.floating_zones.as_deref().unwrap_or(&[]);
    let paragraph_y_offset = input.paragraph_y_offset.unwrap_or(0.0);
    input::validate_float_context(zones, paragraph_y_offset)?;

    if runs.is_empty() {
        if attrs.is_some_and(|a| a.suppress_empty_paragraph_height) {
            return Ok(ParagraphExtentOut {
                kind: "paragraph",
                lines: vec![zero_row()],
                total_height: 0.0,
            });
        }
        let size_pt = attrs
            .and_then(|a| a.default_font_size)
            .unwrap_or(input.defaults.font_size);
        input::validate_pt_size(size_pt, "attrs.defaultFontSize")?;
        let family = attrs
            .and_then(|a| a.default_font_family.as_deref())
            .unwrap_or(&input.defaults.font_family);
        // Empty paragraphs use the regular face.
        let font = regular_chain_head(store, input, family)?;
        return line_filler::empty_paragraph_extent(store, font, size_pt, spacing, &input.compat);
    }

    // ---- single whitespace-only text run measures like an empty paragraph ----
    if runs.len() == 1 && runs[0].kind == "text" && is_whitespace_only(&runs[0]) {
        let run = &runs[0];
        let size_pt = run
            .font_size
            .or_else(|| attrs.and_then(|a| a.default_font_size))
            .unwrap_or(input.defaults.font_size);
        input::validate_pt_size(size_pt, "run.fontSize")?;
        let family = run
            .font_family
            .as_deref()
            .or_else(|| attrs.and_then(|a| a.default_font_family.as_deref()))
            .unwrap_or(&input.defaults.font_family);
        let font = regular_chain_head(store, input, family)?;
        return line_filler::empty_paragraph_extent(store, font, size_pt, spacing, &input.compat);
    }

    // Visible markers consume width only at zero hanging.
    let marker_inline_width = match attrs {
        Some(a) if a.indent.as_ref().and_then(|i| i.hanging).unwrap_or(0.0) == 0.0 => {
            list_marker::list_marker_inline_width(store, input, a)?
        }
        _ => 0.0,
    };

    let prepared = prepare::prepare_runs(store, input)?;

    // Left and right indents shrink both edges; first-line offset affects only the first line.
    let indent = attrs.and_then(|a| a.indent.as_ref());
    let indent_left = indent.and_then(|i| i.left).unwrap_or(0.0);
    let indent_right = indent.and_then(|i| i.right).unwrap_or(0.0);
    let first_line_offset = indent.and_then(|i| i.first_line).unwrap_or(0.0)
        - indent.and_then(|i| i.hanging).unwrap_or(0.0);
    let body_width = (input.max_width - indent_left - indent_right).max(1.0);
    let first_line_width = (body_width - first_line_offset - marker_inline_width).max(1.0);

    line_filler::fill(line_filler::FillParams {
        store,
        prepared: &prepared,
        spacing,
        body_width,
        first_line_width,
        default_font_size_pt: input.defaults.font_size,
        compat: &input.compat,
        tabs: attrs.and_then(|a| a.tabs.as_deref()).unwrap_or(&[]),
        indent_left_px: indent_left,
        first_line_offset_px: first_line_offset,
        zones,
        paragraph_y_offset,
        authoritative_shaping: input.authoritative_shaping,
    })
}

/// JSON boundary: a [`MeasureInput`] envelope in, a serialized
/// [`ParagraphExtentOut`] out. An `Err` starting with `"UNSUPPORTED"` means
/// the host must measure this block with the browser; one starting with
/// `"invalid: "` means the envelope itself was unusable.
pub fn measure_paragraph_json(store: &FontStore, input: &str) -> Result<String, String> {
    let parsed: MeasureInput =
        serde_json::from_str(input).map_err(|e| format!("invalid: parse: {e}"))?;
    let extent = measure_paragraph(store, &parsed).map_err(|e| e.to_string())?;
    serde_json::to_string(&extent).map_err(|e| format!("invalid: serialize: {e}"))
}

fn zero_row() -> TypesetRowOut {
    TypesetRowOut {
        head_run: 0,
        head_char: 0,
        tail_run: 0,
        tail_char: 0,
        width: 0.0,
        ascent: 0.0,
        descent: 0.0,
        line_height: 0.0,
        left_offset: None,
        right_offset: None,
        segments: None,
        float_skip_before: None,
        run_advances: None,
        cluster_advances: None,
        bidi_slices: None,
    }
}

/// Tests for empty or whitespace-only text, including nonbreaking spaces.
fn is_whitespace_only(run: &RunIn) -> bool {
    match run.text.as_deref() {
        None => true,
        Some(t) => t.chars().all(|c| c == '\u{00a0}' || c.is_whitespace()),
    }
}

/// Head of the regular (`|0|0`) chain for `family` — the metrics source for
/// the empty-paragraph and list-marker paths, which never apply bold or
/// italic.
fn regular_chain_head(
    store: &FontStore,
    input: &MeasureInput,
    family: &str,
) -> Result<FontId, MeasureError> {
    let chain = input.chain_for(family, false, false)?;
    prepare::validate_chain(store, &chain)?;
    Ok(chain[0])
}

#[cfg(test)]
mod authoritative_tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn authoritative_json_uses_one_advance_source_for_rows_runs_clusters_and_bidi() {
        let mut store = FontStore::new();
        store.register(FIXTURE.to_vec()).unwrap();
        let input: MeasureInput = serde_json::from_value(serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{
                    "kind": "text",
                    "text": "Latin e\u{301} ffi אב",
                    "letterSpacing": 1.25,
                    "allCaps": false,
                    "kerningMinPt": 14.0
                }]
            },
            "maxWidth": 1000.0,
            "fontChains": { "liberation sans|0|0": [0] },
            "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" },
            "authoritativeShaping": true
        }))
        .unwrap();
        let extent = measure_paragraph(&store, &input).unwrap();
        let line = &extent.lines[0];
        let clusters = line.cluster_advances.as_ref().unwrap();
        let runs = line.run_advances.as_ref().unwrap();
        let slices = line.bidi_slices.as_ref().unwrap();
        let cluster_sum: f32 = clusters.iter().map(|cluster| cluster.advance).sum();
        let run_sum: f32 = runs.iter().map(|run| run.advance).sum();
        let slice_sum: f32 = slices.iter().map(|slice| slice.advance).sum();
        assert!((cluster_sum - line.width).abs() < 0.001);
        assert!((run_sum - line.width).abs() < 0.001);
        assert!((slice_sum - line.width).abs() < 0.001);
        assert!(
            clusters
                .iter()
                .any(|cluster| cluster.end_char - cluster.start_char > 1)
        );
        assert!(slices.iter().any(|slice| slice.bidi_level % 2 == 1));
    }

    #[test]
    fn rotated_inline_image_uses_transformed_footprint_for_flow() {
        let store = FontStore::new();
        let input: MeasureInput = serde_json::from_value(serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{
                    "kind": "image",
                    "width": 80.0,
                    "height": 20.0,
                    "rotationBounds": { "width": 20.0, "height": 80.0 }
                }]
            },
            "maxWidth": 200.0,
            "defaults": { "fontSize": 12.0, "fontFamily": "Fallback" },
            "authoritativeShaping": true
        }))
        .unwrap();
        let extent = measure_paragraph(&store, &input).unwrap();
        assert!((extent.lines[0].width - 20.0).abs() < 0.001);
        assert!(extent.lines[0].line_height >= 80.0);
    }
}
