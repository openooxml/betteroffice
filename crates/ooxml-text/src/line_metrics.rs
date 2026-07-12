use crate::font_store::FontMetrics;
use crate::shape::ShapeFeature;

/// Line-metric compatibility flags.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CompatFlags {
    /// Drop external leading from the font-unit line height.
    pub no_leading: bool,
    /// Do not justify lines ending in a soft return.
    pub do_not_expand_shift_return: bool,
}

/// Line-spacing rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacingRule {
    /// Automatic spacing in 240ths of a line.
    Auto { line_240ths: u32 },
    /// Fixed line box in pixels.
    Exact { px: f32 },
    /// Minimum line box in pixels.
    AtLeast { px: f32 },
}

/// One line box in px. total height = ascent + descent + leading.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct LineBox {
    pub ascent: f32,
    pub descent: f32,
    /// External leading; zero when disabled.
    pub leading: f32,
}

impl LineBox {
    pub fn height(&self) -> f32 {
        self.ascent + self.descent + self.leading
    }
}

/// Compute a single-spaced line box.
pub fn single_line_box(m: &FontMetrics, size_px: f32, compat: &CompatFlags) -> LineBox {
    if m.units_per_em == 0 || size_px.is_nan() || size_px <= 0.0 {
        return LineBox {
            ascent: 0.0,
            descent: 0.0,
            leading: 0.0,
        };
    }
    let scale = size_px / m.units_per_em as f32;

    let ascent = m.os2_win_ascent as f32 * scale;
    let descent = m.os2_win_descent as f32 * scale;

    let leading = if compat.no_leading {
        0.0
    } else {
        // i32 arithmetic: i16/u16 sums cannot overflow, and hhea descender
        // is negative by convention (hence the subtraction).
        let hhea_total = m.hhea_ascender as i32 - m.hhea_descender as i32 + m.hhea_line_gap as i32;
        let win_total = m.os2_win_ascent as i32 + m.os2_win_descent as i32;
        (hhea_total - win_total).max(0) as f32 * scale
    };

    LineBox {
        ascent,
        descent,
        leading,
    }
}

/// Apply a spacing rule to a measured line box.
pub fn apply_spacing_rule(content: LineBox, rule: &LineSpacingRule) -> LineBox {
    match *rule {
        LineSpacingRule::Auto { line_240ths } => {
            let target = content.height() * (line_240ths as f32 / 240.0);
            let core = content.ascent + content.descent;
            if target >= core {
                LineBox {
                    ascent: content.ascent,
                    descent: content.descent,
                    leading: target - core,
                }
            } else {
                let scale = if core > 0.0 { target / core } else { 0.0 };
                LineBox {
                    ascent: content.ascent * scale,
                    descent: content.descent * scale,
                    leading: 0.0,
                }
            }
        }
        LineSpacingRule::Exact { px } => {
            let px = px.max(0.0);
            let descent = content.descent.min(px);
            LineBox {
                ascent: px - descent,
                descent,
                leading: 0.0,
            }
        }
        LineSpacingRule::AtLeast { px } => {
            if content.height() >= px {
                content
            } else {
                LineBox {
                    leading: content.leading + (px - content.height()),
                    ..content
                }
            }
        }
    }
}

pub fn stretch_spaces(advances: &mut [f32], is_space: &[bool], slack: f32) {
    // the explicit NaN test makes a NaN slack a no-op instead of poisoning
    // every space advance
    if slack.is_nan() || slack <= 0.0 {
        return;
    }
    let spaces = advances
        .iter()
        .zip(is_space)
        .filter(|&(_, &space)| space)
        .count();
    if spaces == 0 {
        return;
    }
    let share = slack / spaces as f32;
    for (advance, &space) in advances.iter_mut().zip(is_space) {
        if space {
            *advance += share;
        }
    }
}

pub fn line_is_justified(
    last_line_of_paragraph: bool,
    ends_with_soft_return: bool,
    compat: &CompatFlags,
) -> bool {
    if ends_with_soft_return {
        return !compat.do_not_expand_shift_return;
    }
    !last_line_of_paragraph
}

/// Test whether pair kerning meets its size threshold.
pub fn kern_enabled(font_size_half_points: u32, kern_threshold_half_points: u32) -> bool {
    kern_threshold_half_points != 0 && font_size_half_points >= kern_threshold_half_points
}

/// Feature list to hand [`crate::shape::shape`] for a run whose kerning gate
/// is `enabled` (from [`kern_enabled`]).
///
/// Contract: `enabled == true` returns the empty list — rustybuzz's default
/// features already apply GPOS pair kerning. `enabled == false` returns
/// `kern=0`, which rustybuzz honors even when kerning rides the GPOS `kern`
/// feature of a modern font (proven against the Liberation Sans fixture:
/// `kern=0` shaping of "AV" equals the sum of the pair's hmtx advances).
/// Callers with their own feature lists append these on top.
pub fn kern_features(enabled: bool) -> Vec<ShapeFeature> {
    if enabled {
        Vec::new()
    } else {
        vec![ShapeFeature {
            tag: *b"kern",
            value: 0,
        }]
    }
}
