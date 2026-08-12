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
//! ## 1a. GDI-compatible quantization — [`CompatFlags::gdi_line_metrics`]
//!
//! Word's text stack descends from GDI, whose per-font vertical metrics
//! (`TEXTMETRIC`: `tmAscent`, `tmDescent`, `tmExternalLeading`) are whole
//! logical units, not fractions — DirectWrite still exposes the same
//! quantization as *GDI-compatible metrics* for a given em size.
//! [`CompatFlags::gdi_line_metrics`] reproduces it in three steps:
//!
//! 1. the em size itself snaps to an integer ppem (`round(size_px)`; our
//!    layout unit is already px at 96 DPI, so this is the layout grid);
//! 2. the metric family is chosen — `OS/2` fsSelection **bit 7**
//!    (USE_TYPO_METRICS) hands line spacing to sTypoAscender /
//!    sTypoDescender / sTypoLineGap, otherwise the win/hhea pair above
//!    governs (a font setting bit 7 measures several percent tall without
//!    this, because its `usWin*` box is deliberately the clipping box);
//! 3. ascent, descent and leading are each rounded to whole pixels at that
//!    ppem **before** they are summed. Rounding the sum is not the same
//!    number as summing the rounded parts, and the integer stack rounds the
//!    parts — the difference is the per-line error that accumulates into a
//!    pagination shift down a long document.
//!
//! Off by default: whether Word gates this on `w:compatibilityMode` is
//! unmeasured, so the flag exists for an A/B harness rather than as a
//! document-level policy. `w:spacing` line rules apply to the quantized box
//! ([`apply_spacing_rule`] runs after, never before).
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
//!   the line). The baseline is placed preserving the content descent from
//!   the bottom of the fixed box, matching Word: shrink eats the ascent
//!   side first.
//! - `atLeast`: a floor — the measured content height wins when larger;
//!   when the floor wins, the extra space is leading below the descent.
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

use crate::font_store::FontMetrics;
use crate::shape::ShapeFeature;

/// Compat flags parsed host-side from settings.xml (w:compat, ECMA-376 §17.15.3).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CompatFlags {
    /// w:noLeading — drop external leading from the font-unit line height.
    pub no_leading: bool,
    /// w:doNotExpandShiftReturn — lines ended by a soft return are NOT justified.
    pub do_not_expand_shift_return: bool,
    /// Quantize the single-spacing box to whole pixels at an integer ppem,
    /// honoring `OS/2` USE_TYPO_METRICS (rule 1a). Not a settings.xml flag:
    /// no document parses to it yet, and `false` keeps the float scaling.
    pub gdi_line_metrics: bool,
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

/// Word single-spacing line box for a font at `size_px` (rule 1).
///
/// - `ascent` / `descent` come from `OS/2` usWinAscent / usWinDescent (the
///   GDI `tmHeight` lineage Word uses), scaled by `size_px / units_per_em`.
/// - `leading` is GDI `tmExternalLeading`: `max(0, hhea(ascender − descender
///   + lineGap) − (usWinAscent + usWinDescent))` scaled, placed below the
///   descent. Dropped entirely under [`CompatFlags::no_leading`].
/// - Under [`CompatFlags::gdi_line_metrics`] the box is instead quantized
///   per rule 1a: integer ppem, USE_TYPO_METRICS family selection, and each
///   component rounded to whole pixels before the three are summed.
///
/// Panic-free on malformed metrics: a zero `units_per_em` (or a NaN /
/// non-positive `size_px`) yields an all-zero box rather than NaN/negative
/// geometry — font bytes are attacker-controlled and a degenerate line box
/// is the safe downstream value.
pub fn single_line_box(m: &FontMetrics, size_px: f32, compat: &CompatFlags) -> LineBox {
    if m.units_per_em == 0 || size_px.is_nan() || size_px <= 0.0 {
        return LineBox {
            ascent: 0.0,
            descent: 0.0,
            leading: 0.0,
        };
    }
    if compat.gdi_line_metrics {
        return gdi_line_box(m, size_px, compat.no_leading);
    }
    let scale = size_px / m.units_per_em as f32;

    let ascent = m.os2_win_ascent as f32 * scale;
    let descent = m.os2_win_descent as f32 * scale;

    let leading = if compat.no_leading {
        0.0
    } else {
        win_external_leading(m) as f32 * scale
    };

    LineBox {
        ascent,
        descent,
        leading,
    }
}

/// GDI `tmExternalLeading` in design units: the hhea line height's excess
/// over the win box.
///
/// i32 arithmetic: i16/u16 sums cannot overflow, and hhea descender is
/// negative by convention (hence the subtraction).
fn win_external_leading(m: &FontMetrics) -> i32 {
    let hhea_total = m.hhea_ascender as i32 - m.hhea_descender as i32 + m.hhea_line_gap as i32;
    let win_total = m.os2_win_ascent as i32 + m.os2_win_descent as i32;
    (hhea_total - win_total).max(0)
}

/// Design-space (ascent, descent, leading) of the family that governs line
/// spacing: sTypo* when `OS/2` fsSelection bit 7 asks for it, else win/hhea.
///
/// A font can set the bit and still carry an empty or inverted typo box
/// (absent values, or values that do not sum positive); those fall back to
/// win/hhea rather than collapsing the line to nothing. Components clamp
/// non-negative — a positively-signed sTypoDescender is malformed, not a
/// license to emit negative geometry.
fn line_metric_family(m: &FontMetrics) -> (i32, i32, i32) {
    if m.use_typo_metrics() {
        let ascent = m.os2_typo_ascender as i32;
        let descent = -(m.os2_typo_descender as i32);
        let leading = m.os2_typo_line_gap as i32;
        if ascent + descent + leading > 0 {
            return (ascent.max(0), descent.max(0), leading.max(0));
        }
    }
    (
        m.os2_win_ascent as i32,
        m.os2_win_descent as i32,
        win_external_leading(m),
    )
}

/// Rule 1a: the single-spacing box on GDI's integer grid.
///
/// `no_leading` is applied before quantization, so a dropped leading
/// contributes exactly zero rather than a rounded remainder. f64 keeps the
/// `design × ppem` product exact for every representable design value, so
/// each `round` sees the true ratio and not an accumulated float error.
fn gdi_line_box(m: &FontMetrics, size_px: f32, no_leading: bool) -> LineBox {
    let ppem = (size_px.round() as i32).max(1);
    let (ascent, descent, leading) = line_metric_family(m);
    let leading = if no_leading { 0 } else { leading };

    let upm = m.units_per_em as f64;
    let px = |design: i32| (design as f64 * ppem as f64 / upm).round() as f32;
    LineBox {
        ascent: px(ascent),
        descent: px(descent),
        leading: px(leading),
    }
}

/// Apply `w:spacing` lineRule to a measured content line box (rule 2).
///
/// - `Auto`: target height = `content.height() × line_240ths / 240` — the
///   *full* box including leading is scaled (Word scales line pitch).
///   Ascent/descent are preserved and the delta goes to leading below the
///   descent, so the baseline stays at the top of a taller line box exactly
///   as Word places it (cursor/selection rects hug the text). If the target
///   undercuts ascent + descent (sub-single spacing), ascent and descent
///   shrink proportionally and leading is 0.
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
