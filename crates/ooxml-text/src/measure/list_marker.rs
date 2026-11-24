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

use crate::{FontId, FontStore, ShapedGlyph};

use super::input::{AttrsIn, MeasureRequest};
use super::tabs::twips_to_px;
use super::{MeasureError, pt_to_px};

/// ECMA-376 §17.6.13 default tab stop.
const DEFAULT_TAB_STOP_TWIPS: f32 = 720.0;

#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct InlineListMarker<T> {
    pub width: T,
    pub advance: T,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<InlineListMarkerShape>>,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct InlineListMarkerShape {
    pub family: String,
    pub size_pt: f32,
    pub runs: Vec<InlineListMarkerRun>,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct InlineListMarkerRun {
    pub text: String,
    pub font_id: u32,
    pub level: u8,
    pub glyphs: Vec<ShapedGlyph>,
}

pub(super) fn measure_inline_list_marker(
    store: &FontStore,
    input: &MeasureRequest<'_>,
    max_width: f32,
    allow_metric_fallback: bool,
) -> Result<Option<InlineListMarker<f32>>, MeasureError> {
    let Some(attrs) = input.block.attrs.as_ref() else {
        return Ok(None);
    };
    let hanging = attrs.indent.as_ref().and_then(|i| i.hanging);
    if hanging.unwrap_or(0.0) != 0.0 {
        return Ok(None);
    }
    let Some(marker) = visible_marker(attrs) else {
        return Ok(None);
    };
    let (family, size_pt) = marker_face(input, attrs);
    super::input::validate_pt_size(size_pt, "attrs.listMarkerFontSize")?;
    let size_px = pt_to_px(size_pt);
    let mut runs = Vec::new();
    let width = marker_text_width(store, input, attrs, marker, |text, font, level, glyphs| {
        runs.push(InlineListMarkerRun {
            text: text.to_owned(),
            font_id: store.canonical_face(font).to_u32(),
            level,
            glyphs,
        });
    });
    let width = match width {
        Ok((width, _)) => width,
        Err(_) if allow_metric_fallback => {
            runs.clear();
            marker.chars().count() as f32 * size_px
        }
        Err(error) => return Err(error),
    };
    let space = if attrs.list_marker_suffix.as_deref() == Some("space") {
        match marker_text_width(store, input, attrs, " ", |_, _, _, _| {}) {
            Ok((width, _)) => width,
            Err(_) if allow_metric_fallback => size_px,
            Err(error) => return Err(error),
        }
    } else {
        0.0
    };
    let mut marker = inline_list_marker(attrs, width, space, size_px, max_width)?;
    marker.shape = Some(Box::new(InlineListMarkerShape {
        family: family.to_owned(),
        size_pt,
        runs,
    }));
    Ok(Some(marker))
}

pub fn inline_list_marker(
    attrs: &AttrsIn,
    width: f32,
    space: f32,
    size_px: f32,
    max_width: f32,
) -> Result<InlineListMarker<f32>, MeasureError> {
    let advance = match attrs.list_marker_suffix.as_deref() {
        Some("nothing") => width,
        Some("space") => width + space,
        _ => {
            let indent = attrs.indent.as_ref();
            let start = indent.and_then(|i| i.left).unwrap_or(0.0)
                + indent.and_then(|i| i.first_line).unwrap_or(0.0);
            let available = max_width - start - indent.and_then(|i| i.right).unwrap_or(0.0);
            let gap = (size_px * 0.5).min((available - width - size_px).max(0.0));
            marker_tab_stop(attrs, start + width, false)?
                .map(|stop| stop - start)
                .filter(|&advance| advance <= available - 1.0)
                .unwrap_or(width + gap)
        }
    };
    Ok(InlineListMarker {
        width,
        advance,
        shape: None,
    })
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
    let (natural_width, size_px) =
        marker_text_width(store, input, attrs, marker, |_, _, _, _| {})?;
    let marker_end_px = indent_left - hanging + natural_width;
    if marker_end_px <= indent_left {
        return Ok(0.0);
    }
    let body_start =
        marker_tab_stop(attrs, marker_end_px, true)?.unwrap_or(marker_end_px + size_px * 0.5);
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

fn marker_face<'a>(input: &'a MeasureRequest<'_>, attrs: &'a AttrsIn) -> (&'a str, f32) {
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
    (family, size_pt)
}

/// Measures marker text and returns its width and font size in pixels.
fn marker_text_width(
    store: &FontStore,
    input: &MeasureRequest<'_>,
    attrs: &AttrsIn,
    text: &str,
    shaped_run: impl FnMut(&str, FontId, u8, Vec<ShapedGlyph>),
) -> Result<(f32, f32), MeasureError> {
    let (family, size_pt) = marker_face(input, attrs);
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
    let natural_width = super::prepare::measure_plain_text(
        store,
        &chain,
        text,
        size_px,
        base,
        input.font_chains,
        (attrs.list_marker_bold, attrs.list_marker_italic, family),
        &[],
        shaped_run,
    )?;
    Ok((natural_width, size_px))
}

/// Finds the closest marker tab stop, or none when no stops exist.
fn marker_tab_stop(
    attrs: &AttrsIn,
    marker_end_px: f32,
    include_equal: bool,
) -> Result<Option<f32>, MeasureError> {
    let first_custom_past = attrs
        .tabs
        .as_deref()
        .unwrap_or(&[])
        .iter()
        .filter(|t| t.val != "clear" && t.val != "bar")
        .map(|t| twips_to_px(t.pos))
        .filter(|&px| px > marker_end_px || (include_equal && px == marker_end_px))
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
