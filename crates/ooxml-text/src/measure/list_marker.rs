//! List-marker inline width: how much of the first line a visible marker
//! consumes.
//!
//! The marker's face resolves in precedence order — numbering level, then the
//! paragraph's first text run, then the paragraph default, then the document
//! default — using the marker's bold and italic style. `w:suff` then fixes the footprint:
//! `nothing` is the marker's own width, `space` adds one space glyph, and
//! `tab` (the default) grows the marker out to the nearest stop past its end,
//! taking whichever is closer of the first custom stop and the first
//! `w:defaultTabStop` grid line. With no grid at all — a zero default stop
//! and no custom stops — the marker takes a half-em gap instead.

use crate::font_store::FontStore;

use super::input::{AttrsIn, MeasureRequest};
use super::tabs::twips_to_px;
use super::{MeasureError, pt_to_px};

/// ECMA-376 §17.6.13 default tab stop.
const DEFAULT_TAB_STOP_TWIPS: f32 = 720.0;

/// Marker footprint in pixels. Zero when the paragraph has no marker text or
/// the marker is hidden; callers apply it only at zero hanging indent.
pub(super) fn list_marker_inline_width(
    store: &FontStore,
    input: &MeasureRequest<'_>,
    attrs: &AttrsIn,
) -> Result<f32, MeasureError> {
    let Some(marker) = visible_marker(attrs) else {
        return Ok(0.0);
    };
    let (natural_width, size_px) = marker_text_width(store, input, attrs, marker)?;

    match attrs.list_marker_suffix.as_deref() {
        Some("nothing") => return Ok(natural_width),
        Some("space") => {
            let (space, _) = marker_text_width(store, input, attrs, " ")?;
            return Ok(natural_width + space);
        }
        _ => {}
    }

    let indent = attrs.indent.as_ref();
    let indent_left = indent.and_then(|i| i.left).unwrap_or(0.0);
    let first_line = indent.and_then(|i| i.first_line).unwrap_or(0.0);
    let marker_start_px = indent_left + first_line;
    let marker_end_px = marker_start_px + natural_width;
    // Default suffix `tab`: body text aligns at the closest stop past the marker end.
    match marker_tab_stop(attrs, marker_end_px)? {
        Some(body_start) => Ok(body_start - marker_start_px),
        // No tab grid at all: half-em visual gap after the marker.
        None => Ok(natural_width + size_px * 0.5),
    }
}

/// Extra first-line indent when a tab-suffixed marker overruns its hanging slot.
pub(super) fn list_marker_tab_overrun(
    store: &FontStore,
    input: &MeasureRequest<'_>,
    attrs: &AttrsIn,
) -> Result<f32, MeasureError> {
    if attrs.bidi || !matches!(attrs.list_marker_suffix.as_deref(), None | Some("tab")) {
        return Ok(0.0);
    }
    let indent = attrs.indent.as_ref();
    let indent_left = indent.and_then(|i| i.left).unwrap_or(0.0);
    let hanging = indent.and_then(|i| i.hanging).unwrap_or(0.0);
    if indent_left <= 0.0 || hanging <= 0.0 || hanging > indent_left {
        return Ok(0.0);
    }
    let Some(marker) = visible_marker(attrs) else {
        return Ok(0.0);
    };
    let (natural_width, size_px) = marker_text_width(store, input, attrs, marker)?;
    let marker_end_px = indent_left - hanging + natural_width;
    if marker_end_px <= indent_left {
        return Ok(0.0);
    }
    let body_start =
        marker_tab_stop(attrs, marker_end_px)?.unwrap_or(marker_end_px + size_px * 0.5);
    let overrun = body_start - indent_left;
    let indent_right = indent.and_then(|i| i.right).unwrap_or(0.0);
    let body_width = input.max_width - indent_left - indent_right;
    // Keep today's text start when the stop leaves less than an em for the text.
    Ok(if overrun <= body_width - size_px {
        overrun
    } else {
        0.0
    })
}

/// Returns nonempty marker text only when the marker is visible.
fn visible_marker(attrs: &AttrsIn) -> Option<&str> {
    attrs
        .list_marker
        .as_deref()
        .filter(|marker| !attrs.list_marker_hidden && !marker.is_empty())
}

/// Measures marker text and returns its width and font size in pixels.
fn marker_text_width(
    store: &FontStore,
    input: &MeasureRequest<'_>,
    attrs: &AttrsIn,
    text: &str,
) -> Result<(f32, f32), MeasureError> {
    // Font precedence: level, first text run, paragraph, document.
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

    let chain = input.chain_for(family, attrs.list_marker_bold, attrs.list_marker_italic)?;
    super::prepare::validate_chain(store, &chain)?;
    // Marker text under the paragraph's base direction (`w:bidi` → RTL);
    // like everywhere else, direction affects segmentation, never the sum.
    let base = if attrs.bidi {
        crate::bidi::BaseDirection::Rtl
    } else {
        crate::bidi::BaseDirection::Ltr
    };
    let natural_width = super::prepare::measure_plain_text(store, &chain, text, size_px, base)?;
    Ok((natural_width, size_px))
}

/// Finds the closest marker tab stop, or none when no stops exist.
fn marker_tab_stop(attrs: &AttrsIn, marker_end_px: f32) -> Result<Option<f32>, MeasureError> {
    let first_custom_past = attrs
        .tabs
        .as_deref()
        .unwrap_or(&[])
        .iter()
        .filter(|t| t.val != "clear" && t.val != "bar")
        .map(|t| twips_to_px(t.pos))
        .filter(|&px| px >= marker_end_px)
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
        Some(((marker_end_px / default_tab_stop_px).floor() + 1.0) * default_tab_stop_px)
    } else {
        None
    };

    // Closest wins — a far custom tab must not override a nearer grid stop.
    Ok(match (first_custom_past, first_grid_past) {
        (Some(c), Some(g)) => Some(c.min(g)),
        (c, g) => c.or(g),
    })
}
