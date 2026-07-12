use std::collections::HashMap;

use serde::Deserialize;

use crate::font_store::FontId;

use super::MeasureError;

/// Paragraph measurement request.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MeasureInput {
    pub block: BlockIn,
    pub max_width: f32,
    /// `"<family lowercase>|<b 0|1>|<i 0|1>"` → ordered `FontStore` ids.
    #[serde(default)]
    pub font_chains: HashMap<String, Vec<u32>>,
    pub defaults: DefaultsIn,
    #[serde(default)]
    pub compat: CompatIn,
    #[serde(default)]
    pub floating_zones: Option<Vec<FloatZoneIn>>,
    #[serde(default)]
    pub paragraph_y_offset: Option<f32>,
    #[serde(default)]
    pub authoritative_shaping: bool,
}

impl MeasureInput {
    /// Look up the fallback chain for a `(family, bold, italic)` combination.
    pub(super) fn chain_for(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
    ) -> Result<Vec<FontId>, MeasureError> {
        let key = format!(
            "{}|{}|{}",
            family.to_lowercase(),
            u8::from(bold),
            u8::from(italic)
        );
        let ids = self
            .font_chains
            .get(&key)
            .ok_or_else(|| MeasureError::Unsupported(format!("no font chain for key {key:?}")))?;
        if ids.is_empty() {
            return Err(MeasureError::Unsupported(format!(
                "empty font chain for key {key:?}"
            )));
        }
        Ok(ids.iter().map(|&id| FontId(id)).collect())
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BlockIn {
    pub kind: String,
    #[serde(default)]
    pub runs: Vec<RunIn>,
    #[serde(default)]
    pub attrs: Option<AttrsIn>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DefaultsIn {
    /// Points.
    pub font_size: f32,
    pub font_family: String,
}

/// Line-metric compatibility flags.
#[derive(Debug, Clone, Copy, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CompatIn {
    #[serde(default)]
    pub no_leading: bool,
    #[serde(default)]
    pub do_not_expand_shift_return: bool,
}

/// A measurable inline run.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunIn {
    pub kind: String,
    #[serde(default)]
    pub text: Option<String>,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    #[serde(default)]
    pub bold_cs: Option<bool>,
    #[serde(default)]
    pub italic_cs: Option<bool>,
    #[serde(default)]
    pub font_size: Option<f32>,
    #[serde(default)]
    pub font_size_cs: Option<f32>,
    #[serde(default)]
    pub font_family: Option<String>,
    #[serde(default)]
    pub font_slots: Option<RunFontSlotsIn>,
    #[serde(default)]
    pub complex_script: bool,
    #[serde(default)]
    pub language: Option<RunLanguageSlotsIn>,
    #[serde(default)]
    pub letter_spacing: Option<f32>,
    /// Uppercase before shaping.
    #[serde(default)]
    pub all_caps: bool,
    /// Lowercase chars shape as their uppercase glyph with advances scaled
    /// by the synthesized-small-caps factor (0.7, the Blink/WebKit
    /// multiplier — see `SMALL_CAPS_ADVANCE_SCALE` in `prepare.rs`).
    /// `allCaps` wins when both are set.
    #[serde(default)]
    pub small_caps: bool,
    /// Percent (100 = normal); multiplies glyph advances.
    #[serde(default)]
    pub horizontal_scale: Option<f32>,
    /// Minimum font size in points at which pair kerning is enabled.
    #[serde(default)]
    pub kerning_min_pt: Option<f32>,
    #[serde(default)]
    pub superscript: bool,
    /// See `superscript`.
    #[serde(default)]
    pub subscript: bool,
    /// Hidden/view-suppressed runs retain logical positions but consume no
    /// advance and do not contribute line metrics.
    #[serde(default)]
    pub hidden: bool,
    /// Force right-to-left base direction.
    #[serde(default)]
    pub rtl: bool,
    #[serde(default)]
    pub fallback: Option<String>,
    #[serde(default)]
    pub width: Option<f32>,
    /// Image runs: declared height in px. Missing is treated as zero-size.
    #[serde(default)]
    pub height: Option<f32>,
    /// Post-rotation layout footprint. The original width/height remain the
    /// image content frame used by paint.
    #[serde(default)]
    pub rotation_bounds: Option<RotationBoundsIn>,
    #[serde(default)]
    pub dist_top: Option<f32>,
    #[serde(default)]
    pub dist_bottom: Option<f32>,
    /// Floating-object wrap type.
    #[serde(default)]
    pub wrap_type: Option<String>,
    #[serde(default)]
    pub display_mode: Option<String>,
    #[serde(default)]
    pub position: Option<serde::de::IgnoredAny>,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunFontSlotsIn {
    #[serde(default)]
    pub ascii: Option<String>,
    #[serde(default)]
    pub h_ansi: Option<String>,
    #[serde(default)]
    pub east_asia: Option<String>,
    #[serde(default)]
    pub cs: Option<String>,
    #[serde(default)]
    pub hint: Option<String>,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunLanguageSlotsIn {
    #[serde(default)]
    pub latin: Option<String>,
    #[serde(default)]
    pub east_asia: Option<String>,
    #[serde(default)]
    pub bidi: Option<String>,
}

#[derive(Debug, Clone, Copy, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RotationBoundsIn {
    #[serde(default)]
    pub width: Option<f32>,
    #[serde(default)]
    pub height: Option<f32>,
}

/// Paragraph attributes used during measurement.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AttrsIn {
    #[serde(default)]
    pub alignment: Option<String>,
    #[serde(default)]
    pub spacing: Option<SpacingIn>,
    #[serde(default)]
    pub indent: Option<IndentIn>,
    #[serde(default)]
    pub tabs: Option<Vec<TabStopIn>>,
    /// Force right-to-left paragraph direction.
    #[serde(default)]
    pub bidi: bool,
    /// Points.
    #[serde(default)]
    pub default_font_size: Option<f32>,
    #[serde(default)]
    pub default_font_family: Option<String>,
    /// The "trailing empty paragraph after a table" zero-height anchor.
    #[serde(default)]
    pub suppress_empty_paragraph_height: bool,
    #[serde(default)]
    pub list_marker: Option<String>,
    #[serde(default)]
    pub list_marker_hidden: bool,
    /// Marker font family.
    #[serde(default)]
    pub list_marker_font_family: Option<String>,
    /// Points.
    #[serde(default)]
    pub list_marker_font_size: Option<f32>,
    /// Marker suffix mode.
    #[serde(default)]
    pub list_marker_suffix: Option<String>,
    #[serde(default)]
    pub default_tab_stop_twips: Option<f32>,
}

/// Paragraph spacing.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpacingIn {
    #[serde(default)]
    pub before: Option<f32>,
    #[serde(default)]
    pub after: Option<f32>,
    #[serde(default)]
    pub line: Option<f32>,
    /// `"px"` or `"multiplier"`.
    #[serde(default)]
    pub line_unit: Option<String>,
    /// `"auto"`, `"exact"`, or `"atLeast"`.
    #[serde(default)]
    pub line_rule: Option<String>,
}

impl SpacingIn {
    /// Security clamp on file-derived values: everything finite and within
    /// sane px bounds (multipliers are validated on use in the line filler).
    pub(super) fn validate(&self) -> Result<(), MeasureError> {
        for (name, v) in [
            ("spacing.before", self.before),
            ("spacing.after", self.after),
            ("spacing.line", self.line),
        ] {
            if let Some(v) = v
                && !(v.is_finite() && (-100_000.0..=100_000.0).contains(&v))
            {
                return Err(MeasureError::Unsupported(format!("{name} out of range")));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TabStopIn {
    #[serde(default)]
    pub val: String,
    pub pos: f32,
}

/// Security clamp on the file-derived stop list: bounded count, finite
/// in-range positions (twips).
pub(super) fn validate_tabs(tabs: &[TabStopIn]) -> Result<(), MeasureError> {
    if tabs.len() > super::MAX_TAB_STOPS {
        return Err(MeasureError::Unsupported(format!(
            "too many tab stops ({} > {})",
            tabs.len(),
            super::MAX_TAB_STOPS
        )));
    }
    for stop in tabs {
        if !(stop.pos.is_finite() && stop.pos.abs() <= 1_000_000.0) {
            return Err(MeasureError::Unsupported(
                "tab stop position out of range".to_string(),
            ));
        }
    }
    Ok(())
}

/// `ParagraphIndent` in px.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct IndentIn {
    #[serde(default)]
    pub left: Option<f32>,
    #[serde(default)]
    pub right: Option<f32>,
    #[serde(default)]
    pub first_line: Option<f32>,
    #[serde(default)]
    pub hanging: Option<f32>,
}

impl IndentIn {
    pub(super) fn validate(&self) -> Result<(), MeasureError> {
        for (name, v) in [
            ("indent.left", self.left),
            ("indent.right", self.right),
            ("indent.firstLine", self.first_line),
            ("indent.hanging", self.hanging),
        ] {
            if let Some(v) = v
                && !(v.is_finite() && (-100_000.0..=100_000.0).contains(&v))
            {
                return Err(MeasureError::Unsupported(format!("{name} out of range")));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FloatZoneIn {
    pub left_margin: f32,
    /// Px reserved from the content-area RIGHT edge.
    pub right_margin: f32,
    /// Zone top in px (distTop already subtracted by the extraction).
    pub top_y: f32,
    /// Zone bottom in px (distBottom already added). The Y interval is
    /// half-open on both sides in practice: a line `[top, bottom)` misses
    /// the zone when `lineBottom <= topY` or `lineTop >= bottomY`.
    pub bottom_y: f32,
    #[serde(default)]
    pub segments: Option<Vec<FloatSegmentIn>>,
    /// OOXML `topAndBottom` wrap: a full-width band — no text beside it, any
    /// overlapping line is pushed below.
    #[serde(default)]
    pub full_width_block: bool,
}

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FloatSegmentIn {
    pub left_offset: f32,
    pub available_width: f32,
}

/// Limit for document-relative float offsets.
const MAX_FLOAT_Y_PX: f32 = 10_000_000.0;

pub(super) fn validate_float_context(
    zones: &[FloatZoneIn],
    paragraph_y_offset: f32,
) -> Result<(), MeasureError> {
    if !(paragraph_y_offset.is_finite() && paragraph_y_offset.abs() <= MAX_FLOAT_Y_PX) {
        return Err(MeasureError::Unsupported(
            "paragraphYOffset out of range".to_string(),
        ));
    }
    if zones.len() > super::MAX_FLOAT_ZONES {
        return Err(MeasureError::Unsupported(format!(
            "too many float zones ({} > {})",
            zones.len(),
            super::MAX_FLOAT_ZONES
        )));
    }
    for zone in zones {
        for (name, v) in [
            ("zone.leftMargin", zone.left_margin),
            ("zone.rightMargin", zone.right_margin),
        ] {
            if !(v.is_finite() && (-100_000.0..=100_000.0).contains(&v)) {
                return Err(MeasureError::Unsupported(format!("{name} out of range")));
            }
        }
        for (name, v) in [("zone.topY", zone.top_y), ("zone.bottomY", zone.bottom_y)] {
            if !(v.is_finite() && v.abs() <= MAX_FLOAT_Y_PX) {
                return Err(MeasureError::Unsupported(format!("{name} out of range")));
            }
        }
        let segments = zone.segments.as_deref().unwrap_or(&[]);
        if segments.len() > super::MAX_ZONE_SEGMENTS {
            return Err(MeasureError::Unsupported(format!(
                "too many zone segments ({} > {})",
                segments.len(),
                super::MAX_ZONE_SEGMENTS
            )));
        }
        for seg in segments {
            for (name, v) in [
                ("segment.leftOffset", seg.left_offset),
                ("segment.availableWidth", seg.available_width),
            ] {
                if !(v.is_finite() && (-100_000.0..=100_000.0).contains(&v)) {
                    return Err(MeasureError::Unsupported(format!("{name} out of range")));
                }
            }
        }
    }
    Ok(())
}

/// Maximum accepted font size in points.
pub(super) fn validate_pt_size(size_pt: f32, name: &str) -> Result<(), MeasureError> {
    if size_pt.is_finite() && size_pt > 0.0 && size_pt <= 1638.0 {
        Ok(())
    } else {
        Err(MeasureError::Unsupported(format!("{name} out of range")))
    }
}
