//! Word-specific measurement rules — the places where reproducing Word
//! demands something a generic text engine would not do.
//!
//! Each rule is a free function taking its inputs explicitly, including the
//! compat flags, so nothing here reads global state. ECMA-376 references are
//! to Part 1 (WordprocessingML); element semantics are summarized in
//! `reference/quick-ref/wordprocessingml.md` ("Spacing (w:spacing)" section
//! for line-rule value semantics and the twips/240ths unit table).
//!
//! # 1. Font-unit line height (single spacing) — [`single_line_box`]
//!
//! Word derives the default line height from `OS/2` **usWinAscent +
//! usWinDescent** (the GDI `tmHeight` lineage), *not* from hhea
//! ascender/descender and *not* from sTypo values — which is why
//! [`crate::font_store::FontMetrics`] carries all three families. External
//! leading follows GDI's `tmExternalLeading`:
//!
//! ```text
//! tmExternalLeading = MAX(0, hhea(ascender − descender + lineGap)
//!                            − (usWinAscent + usWinDescent))
//! ```
//!
//! scaled to the requested size, and Word places it *below* the descent
//! (line pitch = ascent + descent + external leading, baseline hugging the
//! top of the pitch). The `w:noLeading` compatibility flag (`w:compat`,
//! ECMA-376 §17.15.3) drops that external leading entirely.
//!
//! # 2. Auto / exact / atLeast spacing — [`apply_spacing_rule`]
//!
//! `w:spacing w:lineRule` (§17.3.1.33):
//!
//! - `auto`: `w:line` is in 240ths of a line (240 = single, 276 = the 1.15
//!   default of recent Word styles, 480 = double). Word scales the *full*
//!   single-spacing pitch — including external leading — by `line/240`.
//!   Ascent and descent stay put; the delta lands in the leading below the
//!   descent, so cursor/selection rects hug the text at the top of the line
//!   box for spacing > single (observable Word behavior). Sub-single values
//!   that undercut ascent+descent shrink ascent/descent proportionally.
//! - `exact`: fixes the line box at the given height regardless of content —
//!   taller glyphs are *clipped* (at render time; measurement never grows
//!   the line). The baseline sits at [`EXACT_BASELINE_RATIO`] of the box, a
//!   constant depending on neither the font nor the size.
//! - `atLeast`: a floor — the measured content height wins when larger;
//!   when the floor wins the slack lands *above* the ascent, so the content
//!   descent is preserved from the bottom of the box.
//!
//! Both fixed rules are measured against Word 16.112. Word quantizes to a
//! 0.25pt device grid; this model is continuous, so an off-grid split differs
//! from Word's raster by up to an eighth of a point.
//!
//! Both fixed rules interact with inline objects (images taller than an
//! exact box also clip).
//!
//! # 3. Justification — [`line_is_justified`], [`stretch_spaces`]
//!
//! `w:jc w:val="both"` (§17.3.1.13) stretches **space clusters only** —
//! never inter-letter gaps — distributing the line's slack in equal shares
//! per expandable space cluster (`"distribute"` is the East Asian variant
//! that does stretch inter-character; not implemented here). The final line
//! of a paragraph is not justified, but a line ended by a soft return
//! (shift-enter, `w:br`) *is* — unless the `w:doNotExpandShiftReturn`
//! compat flag (§17.15.3) restores the non-stretching behavior. The
//! soft-return test takes precedence over the last-line flag. Space stretch
//! happens at line layout, after shaping: shaped cluster advances stay
//! fixed, only space-cluster advances grow.
//!
//! # 4. Snap-to-grid — not applied
//!
//! Where a section defines a document grid (`w:docGrid`, §17.6.5), Word snaps
//! each line's height up to the next grid multiple unless the paragraph or
//! run opts out (`w:snapToGrid` on pPr/rPr, §17.3.1/§17.3.2). Measurement
//! here never snaps: the input carries no grid pitch, so a line's height is
//! whatever rules 1 and 2 compute and nothing more. CJK documents relying on
//! the grid measure slightly short as a result.
//!
//! # 5. Kerning threshold — [`kern_enabled`], [`kern_features`]
//!
//! Word applies pair kerning only when the run's `w:kern` half-point
//! threshold (rPr, §17.3.2) is nonzero and the font size is at or above it.
//! [`mod@crate::shape`] applies default OpenType features (which include GPOS
//! pair kerning via the `kern` feature) unconditionally; callers gate it
//! per run by passing [`kern_features`]`(kern_enabled(..))` as the feature
//! list. rustybuzz honors `kern=0` for GPOS-carried kerning (proven against
//! the Liberation Sans fixture in `tests/ooxml_text.rs`), so no shaping-side
//! switch is needed.
//!
//! # 6. Compatibility flags from `settings.xml` — [`CompatFlags`]
//!
//! `w:compat` / `w:compatSetting` (§17.15.3) select metric eras. The two
//! flags rules 1 and 3 consume — `w:noLeading` and
//! `w:doNotExpandShiftReturn` — are carried by [`CompatFlags`], parsed from
//! `settings.xml` host-side and threaded in as inputs. No rule here reads
//! `compatibilityMode` (12/14/15), `w:useWord97LineBreakRules` or
//! `w:balanceSingleByteDoubleByteWidth`, so a document setting them measures
//! as though they were off.

//!
//! [`CompatFlags::gdi_line_metrics`] and [`CompatFlags::typo_line_spacing`]
//! are independent, opt-in experiments. GDI rounds ppem and components; typo
//! spacing selects version-4 `USE_TYPO_METRICS` with a signed gap. Both remain
//! unavailable to paragraph input because observed Word output did not quantize.

use crate::font_store::FontMetrics;
use crate::shape::ShapeFeature;

/// Compat flags parsed host-side from settings.xml (w:compat, ECMA-376 §17.15.3).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CompatFlags {
    /// w:noLeading — drop external leading from the font-unit line height.
    pub no_leading: bool,
    /// w:doNotExpandShiftReturn — lines ended by a soft return are NOT justified.
    pub do_not_expand_shift_return: bool,
    /// Off-by-default experiment that quantizes ppem and metric components.
    pub gdi_line_metrics: bool,
    /// Off-by-default experiment that selects version-4 `USE_TYPO_METRICS`.
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

/// Fraction of an `exact` line box that sits above the baseline. Constant in
/// Word — neither the font nor the size moves it.
pub const EXACT_BASELINE_RATIO: f32 = 0.8;

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
/// The default path preserves design metrics; experiments bound them to 16 ems.
/// All paths reject degenerate inputs and cap line boxes at Word's 1638pt limit;
/// glyph advances remain uncapped.
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
///   descent, so the baseline stays at the top of a taller line box exactly
///   as Word places it (cursor/selection rects hug the text). Single spacing
///   is an identity. If the target undercuts ascent + descent (sub-single
///   spacing), ascent and descent shrink proportionally and leading is 0.
/// - `Exact`: the box is fixed at `px` regardless of content and split
///   [`EXACT_BASELINE_RATIO`] above the baseline, the rest below. The split
///   is a constant of Word's, not a property of the content, so the box
///   ignores the font entirely; clipping happens at render time.
/// - `AtLeast`: floor — the content box passes through when taller,
///   otherwise the slack goes *above* the ascent and the content descent is
///   preserved from the bottom of the box.
///
/// `Exact` and a floor-active `AtLeast` leave no leading — `ascent + descent
/// == px` exactly — so a consumer centering half-leading and one hanging the
/// baseline off the box top agree. A content-winning `AtLeast` returns the
/// content box untouched, natural leading included.
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
            LineBox {
                ascent: px * EXACT_BASELINE_RATIO,
                descent: px - px * EXACT_BASELINE_RATIO,
                leading: 0.0,
            }
        }
        LineSpacingRule::AtLeast { px } => {
            if content.height() >= px {
                content
            } else {
                LineBox {
                    ascent: px - content.descent,
                    descent: content.descent,
                    leading: 0.0,
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
