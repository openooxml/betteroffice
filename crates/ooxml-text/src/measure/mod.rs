//! Paragraph measurement pipeline — the Rust twin of the TS wrap engine
//! (`packages/core/src/layout/measure/measureParagraph.ts`).
//!
//! Produces, per paragraph, the same `{ kind: "paragraph", lines:
//! TypesetRow[], totalHeight }` shape the TS engine emits from canvas
//! `measureText`, so a TS seam can swap measurement sources block by block.
//! Anything this engine does not cover returns
//! [`MeasureError::Unsupported`] (stringified as `"UNSUPPORTED: <reason>"`)
//! and the host falls back to browser measurement for that block.
//!
//! # What still returns `UNSUPPORTED`
//!
//! Every well-formed, measurable paragraph is now covered — all five run
//! kinds (text, lineBreak, tab, field, image), including inline, floating,
//! and block/`topAndBottom` own-line images. The remaining `UNSUPPORTED`
//! returns are, by category, NOT measurable-content gaps:
//!
//! - **Security clamps / caps** on file-derived numbers and counts (font
//!   size, letter spacing, horizontal scale, spacing, indents, image dims,
//!   tab-stop / zone / segment / run / line counts, over-long run text) —
//!   refusing degenerate input is the point (repo security guidelines: resource limits).
//! - **Host-contract misses**: no/empty font chain for a `(family, bold,
//!   italic)` key the block uses (the host must supply it), and the
//!   `no font in chain covers U+…` backstop (deviation 5 — the font layer's
//!   job; should never fire once coverage is guaranteed).
//! - **Malformed runs**: a mandatory-break control character
//!   (`\n \r \t \v \f \u{85} \u{2028} \u{2029}`) inside a `text` or `field`
//!   run's string. A well-formed DOCX splits those into `lineBreak`/`tab`
//!   runs, so this only guards attacker-crafted or corrupt input.
//! - **Non-paragraph blocks** (`block.kind != "paragraph"`): structural —
//!   tables, images-as-blocks, textboxes, and breaks are measured by other
//!   code paths (`measureBlock`), never routed to `measure_paragraph`.
//!
//! None of these is a measurable-paragraph feature left unimplemented.
//!
//! # Contract mirrored from `measureParagraph.ts`
//!
//! - `TypesetRow` spans: `headRun`/`tailRun` are inclusive run indices;
//!   `headChar` is inclusive and `tailChar` exclusive, both **UTF-16
//!   code-unit indices** into the run's JS string (`resolveLineSegments.ts`
//!   consumes them with `String.prototype.slice`). Rust text is UTF-8, so
//!   every emitted index is converted via `char::len_utf16` sums and always
//!   lands on a `char` boundary — a split surrogate is unrepresentable.
//! - Wrap tolerance `WRAP_SLACK_PX = 0.5`, greedy break at the last
//!   opportunity that fits, TS's fill-then-hard-break behavior for overlong
//!   unbreakable words (minimum one character per line).
//! - Trailing whitespace at a wrap stays on the line it ends and its advance
//!   is **included** in that line's `width` (TS measures whole words
//!   including their trailing space and never subtracts it).
//! - Line height rules mirror `calculateTypographyMetrics`: `exact` /
//!   `atLeast` / `lineUnit: "multiplier" | "px"` / default single spacing,
//!   with the empty-paragraph floor `WORD_SINGLE_LINE_FLOOR = 1.15` applied
//!   for `auto`/`atLeast` rules.
//! - `totalHeight` = sum of line heights **plus** `spacing.before` and
//!   `spacing.after`, exactly like the TS return value.
//!
//! # Documented deviations from `measureParagraph.ts`
//!
//! Each of these is deliberate; the differential harness (design test gate
//! 2) is tolerance-based where they bite.
//!
//! 1. **Metrics come from font tables, not canvas ink.** TS fills
//!    `TypesetRow.ascent/descent` with `actualBoundingBoxAscent/Descent` of
//!    the sample string `"Hg"` and takes `singleLineRatio` from a hardcoded
//!    per-family table (`fontResolver.ts`). Here both derive from the
//!    resolved font's real `OS/2` values: ascent = `size_px ×
//!    usWinAscent/upem`, descent = `size_px × usWinDescent/upem`, and the
//!    single-line basis is their sum — the same formula the TS table was
//!    hand-derived from, without the table's 4-decimal rounding.
//! 2. **Break opportunities are UAX-14** (`crate::line_break`), not the TS
//!    space/hyphen/tab scan — the divergence documented in `line_break.rs`
//!    (CJK-correct, surrogate-safe). Consequences: consecutive spaces glom
//!    into one opportunity, soft hyphens allow a break, and hard-break
//!    prefixes cut at `char` boundaries where TS's UTF-16 binary search
//!    could split a surrogate pair.
//! 3. **`allCaps` uppercases before shaping**, **`smallCaps` shapes
//!    lowercase as uppercase at the 0.7 browser-synthesis scale** (see
//!    `SMALL_CAPS_ADVANCE_SCALE` in `prepare.rs` for the Chromium/WebKit vs
//!    Gecko-0.8 vs Word-≈0.8 decision record), and **`horizontalScale`
//!    multiplies glyph advances**; TS measurement ignores all three (they
//!    are paint-time CSS in `renderParagraph/runs.ts`), so measured widths
//!    there drift from what the painter draws. This engine measures what
//!    will be painted. `letterSpacing` is added per inter-character gap
//!    unscaled.
//! 4. **Widths come from shaping the whole same-font subrange**; TS
//!    re-measures every word / prefix as an isolated string. Kerning across
//!    a hard-break cut (and ligature advances straddling a cut) is therefore
//!    attributed to the pre-cut line here, where TS drops it.
//! 5. **Uncovered characters are refused** (`UNSUPPORTED: no font in chain
//!    covers U+…`) instead of silently falling back to a browser-chosen
//!    font. This is a *backstop*: covering every character is the font
//!    layer's job (the host builds the fallback chain), so once that layer
//!    guarantees the chain always covers the block's text this refusal
//!    should never fire in practice — it is kept only to fail loudly (fall
//!    back to the browser) rather than mismeasure if a gap ever slips
//!    through.
//! 6. **File-derived numbers are clamped/validated** (font size, letter
//!    spacing, horizontal scale, spacing, indents, run/text/line counts) per
//!    the repo security rules; TS trusts them.
//!
//! `letterSpacing` keeps TS's quirk of counting UTF-16 code units: a word of
//! `n` UTF-16 units gets `letterSpacing × (n − 1)` added, so a surrogate
//! pair contributes one internal gap, and no gap is counted between words
//! (TS measures words separately and sums).
//!
//! # Tab, field, image runs and list markers
//!
//! Tab widths mirror the TS tab branch of `measureParagraph` exactly (see
//! `tabs.rs` for the ported grid/width math and its two mirrored
//! simplifications: measurement always uses the 720-twip default grid, and
//! `decimal` stops measure like `start` stops). The following-runs width
//! anchored on `end`/`center` stops sums this engine's shaped run widths, so
//! the documented allCaps / horizontalScale divergences flow into it too —
//! and it counts field fallbacks and image widths (floating images
//! included), like TS `measureInlineWidthAfterTab`.
//!
//! Field runs measure at their `fallback` text (`"1"` when absent/empty)
//! with the run's family/size/bold/italic — no letter spacing or caps,
//! matching the style object the TS field branch builds.
//!
//! Inline images add their declared width to the line advance and grow the
//! line box per TS `finalizeLine`: alone on a line → image height plus the
//! descent buffer above AND below; flowing with text → seated on the
//! baseline (text descent below only). The reserved height is the
//! column-fitted rendered height (painter `max-width: 100%`). Anchored
//! floating images are skipped (absolutely positioned).
//!
//! Block / `topAndBottom` images (`wrapType == "topAndBottom"` or
//! `displayMode == "block"`) take their own line (ported from the non-inline
//! branch of `measureParagraph`): the current line is finished first if it
//! carries content, the image line's footprint is the DECLARED height plus
//! wrap distances (default 6px, not column-fitted — the block painter draws
//! at authored size), the image contributes no width to the line advance, and
//! a fresh line opens after it (a trailing block image emits an empty
//! following line, TS parity). No image run is UNSUPPORTED any more; a
//! dimensionless image (missing width/height) is treated as zero-size rather
//! than refused.
//!
//! A visible list marker on a zero-hanging paragraph narrows the first line
//! by the ported `getListMarkerInlineWidth` footprint (`list_marker.rs`);
//! the host must supply the marker family's regular chain (see `input.rs`).
//!
//! # Float exclusion zones
//!
//! `floatingZones` + `paragraphYOffset` (see `input.rs` for units/origin)
//! mirror the TS float path of `measureParagraph` exactly, via the geometry
//! ports in `floats.rs`:
//!
//! - Per line, the intersecting zones are resolved at the line's cumulative
//!   Y with a **default-font-size single-line estimate** as the probe height
//!   (TS `ptToPx(DEFAULT_FONT_SIZE) × 1.0` — the estimate uses
//!   `defaults.fontSize`, never the line's actual fonts), while cumulative Y
//!   itself advances by each finalized line's *text* typography height —
//!   image-grown line heights and `floatSkipBefore` gaps feed `totalHeight`
//!   but the growth does NOT feed the next line's zone probe (TS
//!   `finalizeLine` adds `typography.lineHeight`, not the image-grown one).
//! - Lines beside a zone shrink by its margins and emit
//!   `leftOffset`/`rightOffset` (omitted when zero, like TS). The tab
//!   branch's content-x includes `leftOffset` exactly like TS
//!   (`lineX = width + leftOffset`).
//! - When the room beside a zone is under `MIN_WRAP_SEGMENT_WIDTH` (24px),
//!   the line hops below the obstruction and the skipped px are emitted as
//!   `floatSkipBefore` on that line (and added to `totalHeight`).
//! - Segment-splitting (centered) zones emit `segments`, ported from TS
//!   `createLineSegments`, including its bails: a multi-run or non-text line
//!   that needs a two-way split emits no segments. The split point is found
//!   at char granularity (never inside a surrogate pair) where TS's binary
//!   search counts UTF-16 units — the hard-break deviation (2) applies.
//!
//! # RTL / bidi
//!
//! Text is split into UBA level runs (`crate::bidi`) before shaping — each
//! shaped segment is a single directional run, and its resolved direction is
//! passed to rustybuzz explicitly. Everything else stays LOGICAL
//! order: line breaking, `headChar`/`tailChar` spans (the TS side slices
//! run text logically in `resolveLineSegments` and lets the painter's
//! `dir`-attributed spans reorder visually), and widths (a line's width is
//! the logical sum of segment advances — direction-independent). The
//! `w:bidi` paragraph flag and per-run `w:rtl` only force the UBA base
//! direction (affecting neutrals), never the sums; TS measurement has no
//! RTL-specific behavior to mirror at all. Hebrew is covered by the
//! vendored fixture and pinned by test; Arabic flows through the same code
//! path (rustybuzz applies joining) but is untested here — no Arabic
//! fixture font is vendored yet. Add it to the differential corpus once a
//! Noto Arabic fixture lands.

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
    /// The input engages a feature this v1 does not cover — the host must
    /// fall back to browser measurement for this block.
    Unsupported(String),
    /// The input violates the measurement contract (malformed JSON, unknown
    /// font ids, non-finite numbers).
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

/// One typeset line — serializes to the TS `TypesetRow` field names.
/// `headChar`/`tailChar` are UTF-16 code-unit indices (see module docs).
/// The four optional float fields are omitted when unset, mirroring TS
/// `finalizeLine`'s conditional assignment (absent, never `null`/0).
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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_advances: Option<Vec<TypesetRunAdvanceOut>>,
    /// Exact shaped cluster advances and visual x offsets.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cluster_advances: Option<Vec<TypesetClusterAdvanceOut>>,
    /// Bidi slices keep logical identity separate from visual paint order.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi_slices: Option<Vec<TypesetBidiSliceOut>>,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRunAdvanceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub logical_order: u32,
}

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

/// One strip of a segment-split line — the TS `TypesetRowSegment` verbatim
/// (`layout/pagination/types.ts`). Spans use the same run/UTF-16 indexing as
/// [`TypesetRowOut`]; `leftOffset`/`availableWidth`/`width` are px.
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

/// Measured paragraph — serializes to the TS `ParagraphExtent` shape.
#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphExtentOut {
    /// Always `"paragraph"`.
    pub kind: &'static str,
    pub lines: Vec<TypesetRowOut>,
    pub total_height: f32,
}

/// Security caps on file-derived counts (repo security guidelines: resource limits).
pub(crate) const MAX_RUNS: usize = 10_000;
pub(crate) const MAX_RUN_TEXT_BYTES: usize = 1_000_000;
pub(crate) const MAX_LINES: usize = 100_000;
pub(crate) const MAX_TAB_STOPS: usize = 1_000;
pub(crate) const MAX_FLOAT_ZONES: usize = 200;
/// TS extraction emits at most 2 strips per zone (`centeredWrapSegments`);
/// the cap only guards a hand-crafted envelope.
pub(crate) const MAX_ZONE_SEGMENTS: usize = 100;

/// points → CSS px (1pt = 96/72 px), the TS `ptToPx`.
pub(crate) fn pt_to_px(pt: f32) -> f32 {
    pt * 96.0 / 72.0
}

/// Measure one paragraph block against `input.maxWidth`.
///
/// Pure and panic-free on untrusted input; every validation failure or
/// uncovered feature comes back as a [`MeasureError`], never a panic.
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

    // ---- empty paragraph (mirrors the TS `runs.length === 0` branch) ----
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
        // TS `calculateEmptyParagraphMetrics` measures the *regular* face
        // (no bold/italic in the style it builds), hence the |0|0 chain.
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

    // A visible list marker eats into the first line's width in TS only when
    // `hanging == 0` (exact `=== 0` — `measureParagraph` zeroes
    // `markerInlineWidth` for any other hanging, positive or negative).
    let marker_inline_width = match attrs {
        Some(a) if a.indent.as_ref().and_then(|i| i.hanging).unwrap_or(0.0) == 0.0 => {
            list_marker::list_marker_inline_width(store, input, a)?
        }
        _ => 0.0,
    };

    let prepared = prepare::prepare_runs(store, input)?;

    // Indent handling, mirroring `measureParagraph`: left/right shrink both
    // edges; firstLineOffset = firstLine − hanging narrows (or widens) the
    // first line only. Float zones adjust these base widths per line inside
    // the filler (TS computes the same pre-float bases first).
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

/// JSON boundary mirroring `docx-layout`'s `layout_to_json` pattern: input
/// JSON in, `ParagraphExtent` JSON out. `Err` strings starting with
/// `"UNSUPPORTED"` mean the host must fall back to browser measurement.
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

/// TS `isEmptyTextRun`: no text, or only whitespace (nbsp counts as space).
fn is_whitespace_only(run: &RunIn) -> bool {
    match run.text.as_deref() {
        None => true,
        Some(t) => t.chars().all(|c| c == '\u{00a0}' || c.is_whitespace()),
    }
}

/// First font of the family's regular (`|0|0`) chain — the face TS's
/// empty-paragraph metrics use.
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
