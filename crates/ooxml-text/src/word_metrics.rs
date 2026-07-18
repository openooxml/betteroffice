//! Word-specific measurement rules.
//!
//! This module makes Rust measurement reproduce *Word*, not a generic text
//! engine. ECMA-376 references are to Part 1 (WordprocessingML); element
//! semantics are summarized in `reference/quick-ref/wordprocessingml.md`
//! ("Spacing (w:spacing)" section for line-rule value semantics and the
//! twips/240ths unit table). Rules are gated by the Word-parity corpus
//! (design test gate 3); the differential harness (gate 2) diffs these
//! results against the browser-measured TS engine, so every deliberate
//! divergence from the TS convention is called out below.
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
//! **Documented divergence from the TS engine**: the TS layout's
//! `singleLineRatio` is `(usWinAscent + usWinDescent) / unitsPerEm` with no
//! external leading at all (`packages/core/src/utils/fontResolver.ts`,
//! `DEFAULT_SINGLE_LINE_RATIO` block; consumed by
//! `calculateTypographyMetrics` in
//! `packages/core/src/layout/measure/measureParagraph.ts`). Rust includes
//! the GDI external-leading term because that is Word's actual line pitch;
//! the differential harness must expect per-line deltas of
//! `tmExternalLeading × size / upem` (Liberation Sans: 67/2048 em) against
//! the TS oracle, and the Word-parity corpus is the tiebreaker.
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
//! compat flag (§17.15.3) restores the non-stretching behavior. This
//! matches the TS painter's gate `!isLastLine || paragraphEndsWithLineBreak`
//! (`packages/core/src/layout/paint/renderParagraph/line.ts`): the
//! soft-return test takes precedence over the last-line flag. Space stretch
//! happens at line layout, after shaping: shaped cluster advances stay
//! fixed, only space-cluster advances grow.
//!
//! # 4. Snap-to-grid — **TODO, not implemented**
//!
//! When a section defines a document grid (`w:docGrid`, §17.6.5), line
//! pitch snaps each line's height up to the next grid multiple unless the
//! paragraph/run opts out (`w:snapToGrid` on pPr/rPr, §17.3.1/§17.3.2).
//! Required for CJK document fidelity; a no-op for the common Western case.
//! Deferred until the CJK slice of the Word-parity corpus lands — the rule
//! needs the section's grid pitch threaded through layout, which does not
//! exist yet.
//!
//! # 5. Kerning threshold — [`kern_enabled`], [`kern_features`]
//!
//! Word applies pair kerning only when the run's `w:kern` half-point
//! threshold (rPr, §17.3.2) is nonzero and the font size is at or above it.
//! [`crate::shape`] applies default OpenType features (which include GPOS
//! pair kerning via the `kern` feature) unconditionally; callers gate it
//! per run by passing [`kern_features`]`(kern_enabled(..))` as the feature
//! list. rustybuzz honors `kern=0` for GPOS-carried kerning (proven against
//! the Liberation Sans fixture in `tests/ooxml_text.rs`), so no shaping-side
//! switch is needed.
//!
//! # 6. Compatibility flags from `settings.xml` — [`CompatFlags`]
//!
//! `w:compat` / `w:compatSetting` (§17.15.3) select metric eras. The flags
//! consumed by rules 1 and 3 (`w:noLeading`, `w:doNotExpandShiftReturn`)
//! are carried by [`CompatFlags`]; they arrive parsed from `settings.xml`
//! on the host side and are threaded into every rule as inputs, not
//! globals. Still to come as the rules that consume them land:
//! `compatibilityMode` (12/14/15), `w:useWord97LineBreakRules` (legacy
//! kinsoku line breaking, layered over [`crate::line_break`]), and
//! `w:balanceSingleByteDoubleByteWidth`.

use crate::font_store::FontMetrics;
use crate::shape::ShapeFeature;

/// Compat flags parsed host-side from settings.xml (w:compat, ECMA-376 §17.15.3).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CompatFlags {
    /// w:noLeading — drop external leading from the font-unit line height.
    pub no_leading: bool,
    /// w:doNotExpandShiftReturn — lines ended by a soft return are NOT justified.
    pub do_not_expand_shift_return: bool,
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

/// One line box in px. total height = ascent + descent + leading.
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
/// Distribution is an **equal share per expandable space cluster** —
/// Word widens every stretchable space by the same amount rather than
/// proportionally to its advance. (This is also what CSS `text-align:
/// justify` does in the TS painter, so the differential harness agrees by
/// construction; the Word-parity corpus is the final arbiter.)
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

/// Rule 3 gating: whether this line participates in justification at all.
/// last_line_of_paragraph is never justified; soft-return lines are justified
/// unless compat.do_not_expand_shift_return.
///
/// The soft-return test takes precedence: a paragraph whose *last* line ends
/// with a `w:br` still stretches (matching the TS painter's
/// `!isLastLine || paragraphEndsWithLineBreak` gate in
/// `packages/core/src/layout/paint/renderParagraph/line.ts`), and under
/// [`CompatFlags::do_not_expand_shift_return`] a *mid-paragraph* soft-return
/// line stops stretching.
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
