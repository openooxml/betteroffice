//! List-marker inline-width resolution — a port of
//! `packages/core/src/layout/measure/listMarkerWidth.ts`
//! (`getListMarkerInlineWidth` + `resolveListMarkerFont`).
//!
//! The painter renders the marker as an inline-block at the start of the
//! first body line; the measurer subtracts the same footprint from the first
//! line's available width. This module mirrors the TS rules:
//!
//! - No marker / hidden marker → 0. (The caller additionally zeroes the
//!   width whenever `indent.hanging != 0`, mirroring the exact `=== 0` guard
//!   in `measureParagraph` — hanging-marker paragraphs indent instead.)
//! - Marker font per §17.9.6: numbering-level rPr (`listMarkerFont*`), else
//!   the first body text run's font, else paragraph/document defaults. The
//!   TS style carries no bold/italic → the family's regular chain.
//! - `w:suff` (§17.9.25): `nothing` → natural width; `space` → natural + one
//!   space glyph; `tab` (default) → grow to the closest stop past the
//!   marker, where custom `attrs.tabs` stops (non-`clear`/`bar`) and the
//!   `defaultTabStopTwips` grid (anchored at 0, NOT at `w:ind`) compete and
//!   the closest wins. `>=` on the custom-stop filter is intentional (a stop
//!   exactly at the marker's right edge is valid).
//! - No grid at all (`defaultTabStopTwips` explicitly 0, no custom stops) →
//!   natural width + half an em.

use crate::font_store::FontStore;

use super::input::{AttrsIn, MeasureInput};
use super::tabs::twips_to_px;
use super::{MeasureError, pt_to_px};

/// TS `DEFAULT_TAB_STOP_TWIPS` (settingsParser): §17.6.13 default.
const DEFAULT_TAB_STOP_TWIPS: f32 = 720.0;

/// The marker's inline-block width in px, or 0 when nothing is rendered.
/// Caller guarantees `hanging == 0` (the only path where TS consumes this).
pub(super) fn list_marker_inline_width(
    store: &FontStore,
    input: &MeasureInput,
    attrs: &AttrsIn,
) -> Result<f32, MeasureError> {
    // TS `!attrs?.listMarker` — empty string is falsy too.
    let marker = match attrs.list_marker.as_deref() {
        Some(m) if !m.is_empty() => m,
        _ => return Ok(0.0),
    };
    if attrs.list_marker_hidden {
        return Ok(0.0);
    }

    // resolveListMarkerFont: level rPr → first text run → paragraph →
    // document defaults.
    let first_text_run = input.block.runs.iter().find(|r| r.kind == "text");
    let family = attrs
        .list_marker_font_family
        .as_deref()
        .or_else(|| first_text_run.and_then(|r| r.font_family.as_deref()))
        .or(attrs.default_font_family.as_deref())
        .unwrap_or(&input.defaults.font_family);
    let size_pt = attrs
        .list_marker_font_size
        .or_else(|| first_text_run.and_then(|r| r.font_size))
        .or(attrs.default_font_size)
        .unwrap_or(input.defaults.font_size);
    super::input::validate_pt_size(size_pt, "attrs.listMarkerFontSize")?;
    let size_px = pt_to_px(size_pt);

    let chain = input.chain_for(family, false, false)?;
    super::prepare::validate_chain(store, &chain)?;
    // Marker text under the paragraph's base direction (`w:bidi` → RTL);
    // like everywhere else, direction affects segmentation, never the sum.
    let base = if attrs.bidi {
        crate::bidi::BaseDirection::Rtl
    } else {
        crate::bidi::BaseDirection::Ltr
    };
    let natural_width = super::prepare::measure_plain_text(store, &chain, marker, size_px, base)?;

    match attrs.list_marker_suffix.as_deref() {
        Some("nothing") => return Ok(natural_width),
        Some("space") => {
            let space = super::prepare::measure_plain_text(store, &chain, " ", size_px, base)?;
            return Ok(natural_width + space);
        }
        _ => {}
    }

    // Default suffix `tab`: body text aligns at the closest stop past
    // `markerStart + naturalWidth`.
    let indent = attrs.indent.as_ref();
    let indent_left = indent.and_then(|i| i.left).unwrap_or(0.0);
    let first_line = indent.and_then(|i| i.first_line).unwrap_or(0.0);
    let marker_start_px = indent_left + first_line;
    let min_body_start = marker_start_px + natural_width;

    let first_custom_past = attrs
        .tabs
        .as_deref()
        .unwrap_or(&[])
        .iter()
        .filter(|t| t.val != "clear" && t.val != "bar")
        .map(|t| twips_to_px(t.pos))
        .filter(|&px| px >= min_body_start)
        .fold(None::<f32>, |acc, px| {
            Some(acc.map_or(px, |best| best.min(px)))
        });

    let default_tab_stop_twips = attrs
        .default_tab_stop_twips
        .unwrap_or(DEFAULT_TAB_STOP_TWIPS);
    if !(default_tab_stop_twips.is_finite() && default_tab_stop_twips.abs() <= 1_000_000.0) {
        return Err(MeasureError::Unsupported(
            "attrs.defaultTabStopTwips out of range".to_string(),
        ));
    }
    let default_tab_stop_px = twips_to_px(default_tab_stop_twips);
    let first_grid_past = if default_tab_stop_px > 0.0 {
        Some(((min_body_start / default_tab_stop_px).floor() + 1.0) * default_tab_stop_px)
    } else {
        None
    };

    // Closest wins — a far custom tab must not override a nearer grid stop.
    let body_start = match (first_custom_past, first_grid_past) {
        (Some(c), Some(g)) => Some(c.min(g)),
        (c, g) => c.or(g),
    };

    match body_start {
        // No tab grid at all: half-em visual gap after the marker.
        None => Ok(natural_width + size_px * 0.5),
        Some(b) => Ok(b - marker_start_px),
    }
}
