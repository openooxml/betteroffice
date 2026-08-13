//! Word-specific line metrics, spacing, justification, and kerning rules.
//!
//! The default line box scales `OS/2` win ascent/descent and hhea external
//! leading as floats. `w:noLeading` drops the leading. [`apply_spacing_rule`]
//! then applies `auto`, `exact`, or `atLeast` spacing.
//!
//! [`CompatFlags::gdi_line_metrics`] and
//! [`CompatFlags::typo_line_spacing`] are independent, opt-in experiments.
//! The first rounds ppem and metric components to whole pixels; the second
//! selects version-4 `USE_TYPO_METRICS`, retaining signed `sTypoLineGap`.
//! Both default to `false`, and paragraph input cannot enable either one.
//!
//! Direct AppleScript/PDF measurements of Word 16.112 rejected metric
//! quantization in compatibility modes 14 and 15. Across four fonts, seven
//! sizes, and 20-line spans, float line pitch matched within ±0.019pt while
//! per-component quantization missed by up to ±0.49pt, with sign varying by
//! size. Across 10–40-character runs, float advances matched within
//! ±0.0145pt while whole-pixel advances missed by up to 0.348pt.
//!
//! Justification stretches space clusters only. [`kern_enabled`] implements
//! the `w:kern` size threshold. Document-grid snapping is not applied because
//! paragraph measurement receives no grid pitch.

use crate::font_store::FontMetrics;
use crate::shape::ShapeFeature;

/// Compat flags parsed host-side from settings.xml (w:compat, ECMA-376 §17.15.3).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CompatFlags {
    /// w:noLeading — drop external leading from the font-unit line height.
    pub no_leading: bool,
    /// w:doNotExpandShiftReturn — lines ended by a soft return are NOT justified.
    pub do_not_expand_shift_return: bool,
    /// Disabled experiment that quantizes ppem and metric components.
    pub gdi_line_metrics: bool,
    /// Disabled experiment that selects version-4 `USE_TYPO_METRICS`.
    pub typo_line_spacing: bool,
}

/// w:spacing lineRule + line value, pre-converted to px by the host where applicable.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacingRule {
    /// lineRule="auto": w:line in 240ths of a line (240 = single, 276 = 1.15, 480 = double).
    Auto { line_240ths: u32 },
    /// lineRule="exact": fixed line box in px; taller content CLIPS.
    Exact { px: f32 },
    /// lineRule="atLeast": floor in px; measured height wins when larger.
    AtLeast { px: f32 },
}

/// One line box in px: total height = ascent + descent + leading. Leading
/// always sits *below* the descent, so the baseline hugs the top of the box.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct LineBox {
    pub ascent: f32,
    pub descent: f32,
    /// External leading distributed per Word's rules; 0.0 under no_leading.
    pub leading: f32,
}

impl LineBox {
    pub fn height(&self) -> f32 {
        self.ascent + self.descent + self.leading
    }
}

/// Word single-spacing line box for a font at `size_px`.
///
/// The default float path preserves every design metric. Opt-in experiment
/// paths bound components to 16 ems. All paths reject degenerate inputs and
/// cap direct callers at Word's 1638pt size limit.
pub fn single_line_box(m: &FontMetrics, size_px: f32, compat: &CompatFlags) -> LineBox {
    if m.units_per_em == 0 || !size_px.is_finite() || size_px <= 0.0 {
        return LineBox {
            ascent: 0.0,
            descent: 0.0,
            leading: 0.0,
        };
    }
    let size_px = size_px.min(MAX_SIZE_PX);

    if !compat.gdi_line_metrics && !compat.typo_line_spacing {
        let scale = size_px / m.units_per_em as f32;
        return LineBox {
            ascent: m.os2_win_ascent as f32 * scale,
            descent: m.os2_win_descent as f32 * scale,
            leading: if compat.no_leading {
                0.0
            } else {
                win_external_leading(m) as f32 * scale
            },
        };
    }

    let (ascent, descent, leading) = experimental_metric_family(m, compat.typo_line_spacing);
    let leading = if compat.no_leading { 0 } else { leading };

    if compat.gdi_line_metrics {
        let ppem = (size_px.round() as i32).max(1) as f64;
        let upm = m.units_per_em as f64;
        let px = |design: i32| (design as f64 * ppem / upm).round() as f32;
        LineBox {
            ascent: px(ascent),
            descent: px(descent),
            leading: px(leading),
        }
    } else {
        let scale = size_px / m.units_per_em as f32;
        LineBox {
            ascent: ascent as f32 * scale,
            descent: descent as f32 * scale,
            leading: leading as f32 * scale,
        }
    }
}

/// Per-component ceiling for the bounded experiments.
const MAX_METRIC_EMS: i32 = 16;

/// Word's 1638pt size limit in px at 96 DPI.
const MAX_SIZE_PX: f32 = 2184.0;

/// hhea line height in excess of the win box, in design units.
fn win_external_leading(m: &FontMetrics) -> i32 {
    let hhea_total = m.hhea_ascender as i32 - m.hhea_descender as i32 + m.hhea_line_gap as i32;
    let win_total = m.os2_win_ascent as i32 + m.os2_win_descent as i32;
    (hhea_total - win_total).max(0)
}

/// Bounded design metrics for the opt-in experiments.
fn experimental_metric_family(m: &FontMetrics, allow_typo: bool) -> (i32, i32, i32) {
    let cap = MAX_METRIC_EMS * m.units_per_em as i32;

    if allow_typo && m.use_typo_metrics() {
        let ascent = m.os2_typo_ascender as i32;
        let descent = -(m.os2_typo_descender as i32);
        let leading = m.os2_typo_line_gap as i32;
        let usable = ascent > 0
            && descent > 0
            && ascent <= cap
            && descent <= cap
            && leading.abs() <= cap
            && ascent + descent + leading > 0;
        if usable {
            return (ascent, descent, leading);
        }
    }
    (
        (m.os2_win_ascent as i32).min(cap),
        (m.os2_win_descent as i32).min(cap),
        win_external_leading(m).min(cap),
    )
}

/// Apply `w:spacing` lineRule to a measured content line box (rule 2).
///
/// - `Auto`: target height = `content.height() × line_240ths / 240` — the
///   *full* box including leading is scaled (Word scales line pitch).
///   Ascent/descent are preserved and the delta goes to leading below the
///   descent. Single spacing is an identity. A sub-single target that
///   undercuts ascent + descent shrinks both proportionally.
/// - `Exact`: the box is fixed at `px` regardless of content; the baseline
///   is placed so the content **descent is preserved bottom-up** (Word's
///   behavior — shrinking eats the ascent side first) and clipping happens
///   at render time. A box smaller than the descent clamps descent to the
///   box and zeroes the ascent; leading is always 0.
/// - `AtLeast`: floor — the content box passes through when taller,
///   otherwise the shortfall is added as leading below the descent.
pub fn apply_spacing_rule(content: LineBox, rule: &LineSpacingRule) -> LineBox {
    match *rule {
        LineSpacingRule::Auto { line_240ths } => {
            if line_240ths == 240 {
                return content;
            }
            let target = content.height() * (line_240ths as f32 / 240.0);
            let core = content.ascent + content.descent;
            if target < core && line_240ths < 240 {
                let scale = if core > 0.0 { target / core } else { 0.0 };
                LineBox {
                    ascent: content.ascent * scale,
                    descent: content.descent * scale,
                    leading: 0.0,
                }
            } else {
                LineBox {
                    ascent: content.ascent,
                    descent: content.descent,
                    leading: target - core,
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

/// Rule 3: distribute `slack` px across space clusters only (never
/// inter-letter). `is_space[i]` marks advance i as an expandable space
/// cluster. No-op when slack <= 0 or no spaces. Mutates advances in place.
///
/// Distribution is an equal share per expandable space cluster.
///
/// Mismatched slice lengths are a caller bug but must not panic: pairing
/// stops at the shorter slice and the excess is left untouched.
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

/// Tests whether a line participates in justification.
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

/// Rule 5: Word applies pair kerning only when rPr w:kern (half-points) is nonzero
/// and the run's font size (half-points) is >= the threshold.
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
