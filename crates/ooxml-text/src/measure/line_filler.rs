//! Greedy line filling — a faithful port of the TS wrap loop in
//! `measureParagraph.ts` (line state, `startNewLine`, `updateMaxFont`,
//! `finalizeLine`, the overlong-word hard-break, and the float-zone
//! per-line margins / skip / segment handling) over the per-character
//! advance tables built by [`super::prepare`].
//!
//! All Word-metrics math routes through the `wm` alias below so the
//! integrator can swap the temporary shim for `crate::word_metrics` by
//! changing that single import.

use crate::font_store::FontId;

use super::floats;
use super::input::{CompatIn, FloatSegmentIn, FloatZoneIn, SpacingIn, TabStopIn};
use super::prepare::{
    CharAdv, PreparedField, PreparedImage, PreparedRun, PreparedTab, PreparedText,
};
use super::{
    MAX_LINES, MeasureError, ParagraphExtentOut, TypesetBidiSliceOut, TypesetClusterAdvanceOut,
    TypesetRowOut, TypesetRowSegmentOut, TypesetRunAdvanceOut, pt_to_px,
};
use crate::word_metrics as wm;

use super::tabs;

/// TS `WRAP_SLACK_PX`: an overshoot under half a CSS px must not force a
/// wrap that Word's exact twip arithmetic would never make.
const WRAP_SLACK_PX: f32 = 0.5;
/// TS `WORD_SINGLE_LINE_FLOOR` for empty paragraphs on auto/atLeast rules.
const WORD_SINGLE_LINE_FLOOR: f32 = 1.15;
/// TS `DEFAULT_SINGLE_LINE_RATIO` (fontResolver) — the single-line basis
/// used when a line carries no font-bearing run (metrics-less fallback).
const DEFAULT_SINGLE_LINE_RATIO: f32 = 1.15;

pub(super) struct FillParams<'a> {
    pub store: &'a crate::font_store::FontStore,
    pub prepared: &'a [PreparedRun],
    pub spacing: Option<&'a SpacingIn>,
    /// Content width for every line after the first (indents applied).
    pub body_width: f32,
    /// Content width for the first line (first-line/hanging offset applied).
    pub first_line_width: f32,
    /// TS `DEFAULT_FONT_SIZE`: seeds each line's max-font tracking.
    pub default_font_size_pt: f32,
    pub compat: &'a CompatIn,
    /// Custom tab stops (`attrs.tabs`), positions in twips.
    pub tabs: &'a [TabStopIn],
    /// `indent.left` in px — tab stops are content-area-relative, so the
    /// indent is added back when converting line x to grid coordinates.
    pub indent_left_px: f32,
    /// `firstLine − hanging` in px, applied to grid x on the first line only.
    pub first_line_offset_px: f32,
    /// Validated float exclusion zones (TS `options.floatingZones`).
    pub zones: &'a [FloatZoneIn],
    /// TS `options.paragraphYOffset`: this paragraph's Y in the zones' space.
    pub paragraph_y_offset: f32,
    pub authoritative_shaping: bool,
}

struct LineState {
    head_run: u32,
    head_char: u32,
    tail_run: u32,
    tail_char: u32,
    width: f32,
    max_font_size_pt: f32,
    max_font: Option<FontId>,
    max_ascent: f32,
    max_descent: f32,
    max_leading: f32,
    /// Tallest inline-image footprint on the line (rendered height + wrap
    /// distances), the TS `maxImageHeightPx`.
    max_image_height_px: f32,
    available: f32,
    /// Float margin from the content left edge (TS `LineState.leftOffset`).
    left_offset: f32,
    /// Float margin from the content right edge.
    right_offset: f32,
    /// Split-segment strips from centered floating exclusions (TS
    /// `segmentZones`); `Some` even when empty, like the TS `?.length` test.
    segment_zones: Option<Vec<FloatSegmentIn>>,
    contributions: Vec<LineContribution>,
}

#[derive(Debug, Clone, Copy)]
struct LineContribution {
    run_index: u32,
    start_char: u32,
    end_char: u32,
    advance: f32,
    level: u8,
    logical_order: u32,
    shaped_cluster: bool,
}

struct Filler<'a> {
    p: &'a FillParams<'a>,
    rule: wm::LineSpacingRule,
    compat: wm::CompatFlags,
    lines: Vec<TypesetRowOut>,
    cur: LineState,
    /// Running Y within the paragraph, probed against the float zones (TS
    /// `cumulativeHeight`). Advances by each line's *text* typography height
    /// — see `finalize_line`.
    cumulative_height: f32,
    /// Skip accrued hopping past floats, attached to the next finalized line
    /// as `floatSkipBefore` (TS `pendingFloatSkip`).
    pending_float_skip: f32,
}

/// TS `skipObstructingFloats` (measureParagraph): when floats leave no
/// usable width (< `MIN_WRAP_SEGMENT_WIDTH`) at the current Y, hop below
/// them; the skipped px accrue on `pending` and add to `cumulative`.
fn skip_obstructing_floats(
    p: &FillParams,
    line_height: f32,
    line_max_width: f32,
    cumulative: &mut f32,
    pending: &mut f32,
) {
    if p.zones.is_empty() {
        return;
    }
    let absolute_y = p.paragraph_y_offset + *cumulative;
    let skip = floats::find_clear_line_y(
        absolute_y,
        line_height,
        p.zones,
        line_max_width,
        floats::MIN_WRAP_SEGMENT_WIDTH,
    ) - absolute_y;
    if skip > 0.0 {
        *cumulative += skip;
        *pending += skip;
    }
}

/// The probe height for zone tests: TS uses `ptToPx(DEFAULT_FONT_SIZE) ×
/// DEFAULT_LINE_HEIGHT_MULTIPLIER` (multiplier 1.0) with its hardcoded 11pt
/// constant — the same value the host passes as `defaults.fontSize` — never
/// the line's actual fonts (metrics aren't known until the line finalizes).
fn estimated_line_height(p: &FillParams) -> f32 {
    pt_to_px(p.default_font_size_pt)
}

/// Wrap the prepared runs into lines and total the paragraph height
/// (including `spacing.before`/`after` and any `floatSkipBefore` gaps,
/// mirroring the TS return value).
pub(super) fn fill(p: FillParams) -> Result<ParagraphExtentOut, MeasureError> {
    let rule = rule_from_spacing(p.spacing);
    let compat = to_flags(p.compat);

    // TS measureParagraph: probe the zones at the paragraph top with the
    // default-size estimate, hop past obstructions, then resolve the first
    // line's margins at the (possibly skipped-to) Y. The skip probe uses the
    // pre-float first-line base width, like the TS call site.
    let mut cumulative_height = 0.0f32;
    let mut pending_float_skip = 0.0f32;
    let estimated = estimated_line_height(&p);
    skip_obstructing_floats(
        &p,
        estimated,
        p.first_line_width,
        &mut cumulative_height,
        &mut pending_float_skip,
    );
    let first_margins =
        floats::floating_margins(cumulative_height, estimated, p.zones, p.paragraph_y_offset);
    let first_available = floats::available_width(&first_margins, p.first_line_width).max(1.0);

    let mut filler = Filler {
        cur: LineState {
            head_run: 0,
            head_char: 0,
            tail_run: 0,
            tail_char: 0,
            width: 0.0,
            max_font_size_pt: p.default_font_size_pt,
            max_font: None,
            max_ascent: 0.0,
            max_descent: 0.0,
            max_leading: 0.0,
            max_image_height_px: 0.0,
            available: first_available,
            left_offset: first_margins.left,
            right_offset: first_margins.right,
            segment_zones: first_margins.segments,
            contributions: Vec::new(),
        },
        p: &p,
        rule,
        compat,
        lines: Vec::new(),
        cumulative_height,
        pending_float_skip,
    };
    filler.run()?;

    // TS: totalHeight = Σ (lineHeight + floatSkipBefore) + spacing.
    let mut total: f32 = filler
        .lines
        .iter()
        .map(|l| l.line_height + l.float_skip_before.unwrap_or(0.0))
        .sum();
    if let Some(sp) = p.spacing {
        total += sp.before.unwrap_or(0.0) + sp.after.unwrap_or(0.0);
    }
    Ok(ParagraphExtentOut {
        kind: "paragraph",
        lines: filler.lines,
        total_height: total,
    })
}

/// Empty / whitespace-only paragraph: one zero-width line at the resolved
/// line height, floored at `fontSizePx × 1.15` for auto/atLeast rules
/// (TS `calculateEmptyParagraphMetrics`).
pub(super) fn empty_paragraph_extent(
    store: &crate::font_store::FontStore,
    font: FontId,
    size_pt: f32,
    spacing: Option<&SpacingIn>,
    compat: &CompatIn,
) -> Result<ParagraphExtentOut, MeasureError> {
    let metrics = store
        .metrics(font)
        .map_err(|e| MeasureError::Invalid(e.to_string()))?;
    let size_px = pt_to_px(size_pt);
    let content = wm::single_line_box(metrics, size_px, &to_flags(compat));
    let ruled = wm::apply_spacing_rule(content, &rule_from_spacing(spacing));
    let mut line_height = ruled.height();
    if floor_applies(spacing) {
        line_height = line_height.max(size_px * WORD_SINGLE_LINE_FLOOR);
    }

    let mut total = line_height;
    if let Some(sp) = spacing {
        total += sp.before.unwrap_or(0.0) + sp.after.unwrap_or(0.0);
    }
    Ok(ParagraphExtentOut {
        kind: "paragraph",
        lines: vec![TypesetRowOut {
            head_run: 0,
            head_char: 0,
            tail_run: 0,
            tail_char: 0,
            width: 0.0,
            ascent: content.ascent,
            descent: content.descent,
            line_height,
            left_offset: None,
            right_offset: None,
            segments: None,
            float_skip_before: None,
            run_advances: None,
            cluster_advances: None,
            bidi_slices: None,
        }],
        total_height: total,
    })
}

impl Filler<'_> {
    fn run(&mut self) -> Result<(), MeasureError> {
        for (run_index, prun) in self.p.prepared.iter().enumerate() {
            let ri = run_index as u32;
            match prun {
                PreparedRun::LineBreak => {
                    // TS: soft return closes the line at char 0 of the break
                    // run and opens the next line after it.
                    self.cur.tail_run = ri;
                    self.cur.tail_char = 0;
                    self.start_new_line(ri + 1, 0)?;
                }
                PreparedRun::Text(t) => self.fill_text_run(ri, t)?,
                PreparedRun::Tab(t) => self.fill_tab_run(run_index, *t)?,
                PreparedRun::Field(f) => self.fill_field_run(ri, *f)?,
                PreparedRun::InlineImage(img) => self.fill_inline_image(ri, *img)?,
                PreparedRun::OwnLineImage(img) => self.fill_own_line_image(ri, *img)?,
                PreparedRun::SkippedImage { bidi_level, .. } => {
                    // Truly floating image: absolutely positioned, no line
                    // contribution — just advance the tail span (TS parity).
                    self.cur.tail_run = ri;
                    self.cur.tail_char = 1;
                    self.record_atomic(ri, 0, 1, 0.0, *bidi_level);
                }
                PreparedRun::Hidden { utf16_len } => {
                    self.cur.tail_run = ri;
                    self.cur.tail_char = *utf16_len;
                }
            }
        }
        self.finalize_line()
    }

    /// TS inline-image branch: the width joins the line advance (no
    /// empty-line guard — TS wraps an oversize image off an empty line too,
    /// emitting an empty row); the footprint recorded for line growth is the
    /// *rendered* height — the painter fits inline images to the column with
    /// `max-width: 100%` — plus the wrap distances. `updateMaxFont` is NOT
    /// called (images carry no font).
    fn fill_inline_image(&mut self, ri: u32, img: PreparedImage) -> Result<(), MeasureError> {
        if self.cur.width + img.width > self.cur.available + WRAP_SLACK_PX {
            self.start_new_line(ri, 0)?;
        }
        let fit_scale = if img.width > 0.0 && img.width > self.cur.available {
            self.cur.available / img.width
        } else {
            1.0
        };
        let footprint = img.height * fit_scale + img.dist_top + img.dist_bottom;
        if footprint > self.cur.max_image_height_px {
            self.cur.max_image_height_px = footprint;
        }
        self.record_atomic(ri, 0, 1, img.width, img.bidi_level);
        self.cur.width += img.width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    /// TS block-image branch (`wrapType === 'topAndBottom' || displayMode ===
    /// 'block'`): the image gets its own line. If the current line already
    /// carries content, finish it first (unconditionally — no column-fit wrap
    /// check, unlike the inline path). The image line's footprint is the
    /// DECLARED image height plus its wrap distances, assigned straight to
    /// `max_image_height_px` (no `max-width` column scaling — the block
    /// painter draws the image at its authored size), and the image adds NO
    /// width to the line advance. A fresh line opens after it for subsequent
    /// content; when the image is the paragraph's last run that trailing line
    /// is closed empty by the post-loop `finalize_line`, exactly like TS.
    ///
    /// `update_max_font` is NOT called (images carry no font), so a lone
    /// own-line image finalizes through the metrics-less fallback and hits
    /// `finalize_line`'s image-alone branch (`head_run == tail_run`).
    fn fill_own_line_image(&mut self, ri: u32, img: PreparedImage) -> Result<(), MeasureError> {
        if self.cur.width > 0.0 {
            self.start_new_line(ri, 0)?;
        }
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        self.cur.max_image_height_px = img.height + img.dist_top + img.dist_bottom;
        self.record_atomic(ri, 0, 1, 0.0, img.bidi_level);
        self.start_new_line(ri + 1, 0)?;
        Ok(())
    }

    /// TS field branch: the pre-measured fallback width flows like one
    /// unbreakable glyph — wrap first if it doesn't fit a non-empty line.
    fn fill_field_run(&mut self, ri: u32, f: PreparedField) -> Result<(), MeasureError> {
        self.update_max_font(f.font_size_pt, f.metrics_font, f.baseline_shift_px);
        if self.cur.width > 0.0 && self.cur.width + f.width > self.cur.available + WRAP_SLACK_PX {
            self.start_new_line(ri, 0)?;
            self.update_max_font(f.font_size_pt, f.metrics_font, f.baseline_shift_px);
        }
        self.record_atomic(ri, 0, 1, f.width, f.bidi_level);
        self.cur.width += f.width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    /// TS tab branch: width from the shared tab-stop model at the line's
    /// current x (content-area coordinates), the following-runs width
    /// anchored on `end`/`center` stops, the TOC-style clamp when the stop
    /// sits past the line edge, then the ordinary wrap check. On a wrap the
    /// pre-wrap tab width is kept — TS does not recompute it for the new
    /// line's x.
    fn fill_tab_run(&mut self, run_index: usize, t: PreparedTab) -> Result<(), MeasureError> {
        let ri = run_index as u32;
        self.update_max_font(t.font_size_pt, t.metrics_font, 0.0);

        let following = self.following_width_after(run_index);
        // TS: `lineX = currentLine.width + (currentLine.leftOffset ?? 0)` —
        // a float's left margin shifts the tab's content-x (and the
        // past-the-edge clamp below) but not the plain wrap check.
        let line_x = self.cur.width + self.cur.left_offset;
        let is_first_line = self.lines.is_empty();
        let content_x = self.p.indent_left_px
            + if is_first_line {
                self.p.first_line_offset_px
            } else {
                0.0
            }
            + line_x;
        let mut tab_width = tabs::calculate_tab_width(
            content_x,
            self.p.tabs,
            tabs::px_to_twips(self.p.indent_left_px),
            following,
        );

        // Tab targeting a position past the line edge (Word's TOC styles
        // author right stops a hair past the margin): snap to the margin and
        // reserve room for the following runs.
        if line_x + tab_width > self.cur.available + WRAP_SLACK_PX {
            let clamped = self.cur.available - line_x - following;
            if clamped > 1.0 {
                tab_width = clamped;
            }
        }

        if self.cur.width + tab_width > self.cur.available + WRAP_SLACK_PX {
            // line already full of preceding content
            self.start_new_line(ri, 0)?;
            self.update_max_font(t.font_size_pt, t.metrics_font, 0.0);
        }

        self.record_atomic(ri, 0, 1, tab_width, t.bidi_level);
        self.cur.width += tab_width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    /// TS `measureInlineWidthAfterTab`: inline widths of the runs after a
    /// tab, up to (not including) the next tab or line break.
    fn following_width_after(&self, tab_index: usize) -> f32 {
        let mut width = 0.0f32;
        for prun in &self.p.prepared[tab_index + 1..] {
            match prun {
                PreparedRun::Tab(_) | PreparedRun::LineBreak => break,
                PreparedRun::Text(t) => width += span_width(&t.chars, t.letter_spacing),
                PreparedRun::Field(f) => width += f.width,
                // TS sums `next.width || 0` for any image run — inline,
                // block/own-line, and floating alike.
                PreparedRun::InlineImage(img) => width += img.width,
                PreparedRun::OwnLineImage(img) => width += img.width,
                PreparedRun::SkippedImage { width: w, .. } => width += w,
                PreparedRun::Hidden { .. } => {}
            }
        }
        width
    }

    fn fill_text_run(&mut self, ri: u32, t: &PreparedText) -> Result<(), MeasureError> {
        // TS calls updateMaxFont before the empty-text check, so even an
        // empty run contributes its font to the line's metrics.
        self.update_max_font(t.font_size_pt, t.metrics_font, t.baseline_shift_px);
        if t.chars.is_empty() {
            self.cur.tail_run = ri;
            self.cur.tail_char = 0;
            return Ok(());
        }

        let mut char_idx = 0usize;
        let mut break_cursor = 0usize;
        while char_idx < t.chars.len() {
            while break_cursor < t.breaks.len() && t.breaks[break_cursor] <= char_idx {
                break_cursor += 1;
            }
            let next_break = t.breaks.get(break_cursor).copied().unwrap_or(t.chars.len());

            // The word includes its trailing space; its full width lands on
            // the line it ends (TypesetRow.width keeps trailing spaces).
            let word = &t.chars[char_idx..next_break];
            let word_width = span_width(word, t.letter_spacing);

            if word_width > self.cur.available + WRAP_SLACK_PX {
                // Overlong unbreakable word: fill the remaining space on the
                // current line, then hard-break char by char, minimum one
                // char per line (TS's findMaxFittingLength loop).
                let mut chunk_start = 0usize;
                while chunk_start < word.len() {
                    let space_left = self.cur.available - self.cur.width + WRAP_SLACK_PX;
                    let remaining = &word[chunk_start..];
                    let mut best = max_fitting(remaining, t.letter_spacing, space_left);
                    if best == 0 {
                        if self.cur.width > 0.0 {
                            self.start_new_line(ri, utf16_at(t, char_idx + chunk_start))?;
                            self.update_max_font(
                                t.font_size_pt,
                                t.metrics_font,
                                t.baseline_shift_px,
                            );
                            continue;
                        }
                        best = 1;
                    }
                    let chunk = &remaining[..best];
                    let chunk_width = span_width(chunk, t.letter_spacing);
                    self.record_text_clusters(ri, chunk, t.letter_spacing);
                    self.cur.width += chunk_width;
                    self.cur.tail_run = ri;
                    self.cur.tail_char = utf16_at(t, char_idx + chunk_start + best);
                    chunk_start += best;
                    if chunk_start < word.len() {
                        self.start_new_line(ri, utf16_at(t, char_idx + chunk_start))?;
                        self.update_max_font(t.font_size_pt, t.metrics_font, t.baseline_shift_px);
                    }
                }
                char_idx = next_break;
                continue;
            }

            if self.cur.width > 0.0
                && self.cur.width + word_width > self.cur.available + WRAP_SLACK_PX
            {
                self.start_new_line(ri, utf16_at(t, char_idx))?;
                self.update_max_font(t.font_size_pt, t.metrics_font, t.baseline_shift_px);
            }

            self.record_text_clusters(ri, word, t.letter_spacing);
            self.cur.width += word_width;
            self.cur.tail_run = ri;
            self.cur.tail_char = utf16_at(t, next_break);
            char_idx = next_break;
        }
        Ok(())
    }

    /// TS `updateMaxFont`: the first font-bearing run claims the line; after
    /// that only a strictly larger size replaces the metrics source.
    fn update_max_font(&mut self, font_size_pt: f32, font: FontId, baseline_shift_px: f32) {
        if self.cur.max_font.is_none() || font_size_pt > self.cur.max_font_size_pt {
            self.cur.max_font_size_pt = font_size_pt;
            self.cur.max_font = Some(font);
        }
        if let Ok(metrics) = self.p.store.metrics(font) {
            let line = wm::single_line_box(metrics, pt_to_px(font_size_pt), &self.compat);
            self.cur.max_ascent = self
                .cur
                .max_ascent
                .max((line.ascent + baseline_shift_px).max(0.0));
            self.cur.max_descent = self
                .cur
                .max_descent
                .max((line.descent - baseline_shift_px).max(0.0));
            self.cur.max_leading = self.cur.max_leading.max(line.leading);
        }
    }

    fn record_atomic(
        &mut self,
        run_index: u32,
        start_char: u32,
        end_char: u32,
        advance: f32,
        level: u8,
    ) {
        self.cur.contributions.push(LineContribution {
            run_index,
            start_char,
            end_char,
            advance,
            level,
            logical_order: run_index.saturating_mul(1_000_000),
            shaped_cluster: false,
        });
    }

    fn record_text_clusters(&mut self, run_index: u32, chars: &[CharAdv], spacing: f32) {
        for (index, cluster) in chars.iter().enumerate() {
            self.update_max_font(
                cluster.font_size_pt,
                cluster.metrics_font,
                cluster.baseline_shift_px,
            );
            self.cur.contributions.push(LineContribution {
                run_index,
                start_char: cluster.utf16_offset,
                end_char: cluster.utf16_offset + cluster.utf16_len,
                advance: cluster.advance
                    + if index + 1 < chars.len() {
                        spacing
                    } else {
                        0.0
                    },
                level: cluster.level,
                logical_order: run_index
                    .saturating_mul(1_000_000)
                    .saturating_add(cluster.logical_order),
                shaped_cluster: true,
            });
        }
    }

    /// TS `finalizeLine`: resolve typography from the line's largest font,
    /// attach float offsets/segments/skip, and push the row.
    fn finalize_line(&mut self) -> Result<(), MeasureError> {
        if self.lines.len() >= MAX_LINES {
            return Err(MeasureError::Unsupported(format!(
                "too many lines (> {MAX_LINES})"
            )));
        }
        let size_px = pt_to_px(self.cur.max_font_size_pt);
        let content = match self.cur.max_font {
            Some(_) => wm::LineBox {
                ascent: self.cur.max_ascent,
                descent: self.cur.max_descent,
                leading: self.cur.max_leading,
            },
            // No font-bearing run on this line — TS's metrics-less fallback:
            // 0.8/0.2 em split, DEFAULT_SINGLE_LINE_RATIO basis.
            None => wm::LineBox {
                ascent: size_px * 0.8,
                descent: size_px * 0.2,
                leading: size_px * (DEFAULT_SINGLE_LINE_RATIO - 1.0),
            },
        };
        let ruled = wm::apply_spacing_rule(content, &self.rule);
        let mut ascent = content.ascent;
        let text_line_height = ruled.height();
        let mut line_height = text_line_height;

        // TS `finalizeLine`: an inline image taller than the ruled text
        // height grows the line. An image alone on the line (headRun ==
        // tailRun) gets the parent font's descent as breathing room on BOTH
        // sides; an image flowing with text seats on the baseline — full
        // image height above, only the text descent below. The descent
        // always stays text metrics. This must stay in sync with the
        // painter's image-only test in `renderLine` (paired strategies).
        if self.cur.max_image_height_px > line_height {
            let image_h = self.cur.max_image_height_px;
            let buffer = content.descent;
            if self.cur.head_run == self.cur.tail_run {
                line_height = image_h + buffer * 2.0;
                ascent = image_h + buffer;
            } else {
                line_height = image_h + buffer;
                ascent = image_h;
            }
        }

        // TS emits the float fields only when set: offsets > 0, a non-empty
        // segment-zone list (whose port may still decline), a pending skip.
        let segments = match self.cur.segment_zones.as_deref() {
            Some(zones) if !zones.is_empty() => self.create_line_segments(zones),
            _ => None,
        };
        let float_skip_before = (self.pending_float_skip > 0.0).then_some(self.pending_float_skip);
        self.pending_float_skip = 0.0;
        let (run_advances, cluster_advances, bidi_slices) = if self.p.authoritative_shaping {
            advance_metadata(&self.cur.contributions)
        } else {
            (None, None, None)
        };

        self.lines.push(TypesetRowOut {
            head_run: self.cur.head_run,
            head_char: self.cur.head_char,
            tail_run: self.cur.tail_run,
            tail_char: self.cur.tail_char,
            width: self.cur.width,
            ascent,
            descent: content.descent,
            line_height,
            left_offset: (self.cur.left_offset > 0.0).then_some(self.cur.left_offset),
            right_offset: (self.cur.right_offset > 0.0).then_some(self.cur.right_offset),
            segments,
            float_skip_before,
            run_advances,
            cluster_advances,
            bidi_slices,
        });

        // TS advances `cumulativeHeight` by `typography.lineHeight` — the
        // TEXT height, not the image-grown one — so the next line's zone
        // probe deliberately ignores image growth (quirk kept for parity).
        self.cumulative_height += text_line_height;
        Ok(())
    }

    /// TS `createLineSegments`: split the just-finalized line across the
    /// zone's strips. One strip (or a line that fits the first strip within
    /// the wrap slack) → a single segment covering the whole line. A two-way
    /// split only applies to a single-text-run line: the split point is the
    /// longest prefix fitting the first strip (TS `findMaxFittingLength`;
    /// here at char granularity — never inside a surrogate pair), and each
    /// side is re-measured like TS's isolated `measureTextWidth` calls.
    /// Anything else — multi-run line, non-text run, empty or degenerate
    /// split — emits no segments, exactly like the TS `undefined` bails.
    fn create_line_segments(
        &self,
        segment_zones: &[FloatSegmentIn],
    ) -> Option<Vec<TypesetRowSegmentOut>> {
        let cur = &self.cur;
        let first = segment_zones.first()?;
        let second = segment_zones.get(1);

        if second.is_none() || cur.width <= first.available_width + WRAP_SLACK_PX {
            return Some(vec![TypesetRowSegmentOut {
                head_run: cur.head_run,
                head_char: cur.head_char,
                tail_run: cur.tail_run,
                tail_char: cur.tail_char,
                left_offset: first.left_offset,
                available_width: first.available_width,
                width: cur.width,
            }]);
        }
        let second = second?;

        if cur.head_run != cur.tail_run {
            return None;
        }
        let PreparedRun::Text(t) = self.p.prepared.get(cur.head_run as usize)? else {
            return None;
        };

        // The line's char span (CharAdv.utf16_offset is absolute in the run,
        // so the TS `headChar + firstLength` arithmetic falls out for free).
        let start = t.chars.partition_point(|c| c.utf16_offset < cur.head_char);
        let end = t.chars.partition_point(|c| c.utf16_offset < cur.tail_char);
        let chars = &t.chars[start..end];

        let best = max_fitting(chars, t.letter_spacing, first.available_width);
        if best == 0 || best >= chars.len() {
            return None;
        }
        let split_char = chars[best].utf16_offset;

        Some(vec![
            TypesetRowSegmentOut {
                head_run: cur.head_run,
                head_char: cur.head_char,
                tail_run: cur.tail_run,
                tail_char: split_char,
                left_offset: first.left_offset,
                available_width: first.available_width,
                width: span_width(&chars[..best], t.letter_spacing),
            },
            TypesetRowSegmentOut {
                head_run: cur.head_run,
                head_char: split_char,
                tail_run: cur.tail_run,
                tail_char: cur.tail_char,
                left_offset: second.left_offset,
                available_width: second.available_width,
                width: span_width(&chars[best..], t.letter_spacing),
            },
        ])
    }

    /// TS `startNewLine`: finalize, then open a fresh line at `(run, char)`
    /// with the body content width — hopped past and narrowed by the float
    /// zones at the new line's Y — and reset font tracking.
    fn start_new_line(&mut self, run: u32, char_utf16: u32) -> Result<(), MeasureError> {
        self.finalize_line()?;
        let estimated = estimated_line_height(self.p);
        skip_obstructing_floats(
            self.p,
            estimated,
            self.p.body_width,
            &mut self.cumulative_height,
            &mut self.pending_float_skip,
        );
        let margins = floats::floating_margins(
            self.cumulative_height,
            estimated,
            self.p.zones,
            self.p.paragraph_y_offset,
        );
        let available = floats::available_width(&margins, self.p.body_width).max(1.0);
        self.cur = LineState {
            head_run: run,
            head_char: char_utf16,
            tail_run: run,
            tail_char: char_utf16,
            width: 0.0,
            max_font_size_pt: self.p.default_font_size_pt,
            max_font: None,
            max_ascent: 0.0,
            max_descent: 0.0,
            max_leading: 0.0,
            max_image_height_px: 0.0,
            available,
            left_offset: margins.left,
            right_offset: margins.right,
            segment_zones: margins.segments,
            contributions: Vec::new(),
        };
        Ok(())
    }
}

fn advance_metadata(
    contributions: &[LineContribution],
) -> (
    Option<Vec<TypesetRunAdvanceOut>>,
    Option<Vec<TypesetClusterAdvanceOut>>,
    Option<Vec<TypesetBidiSliceOut>>,
) {
    if contributions.is_empty() {
        return (None, None, None);
    }
    let levels: Vec<u8> = contributions.iter().map(|c| c.level).collect();
    let visual = crate::bidi::visual_order_for_levels(&levels);
    let mut runs: Vec<TypesetRunAdvanceOut> = Vec::new();
    let mut clusters = Vec::new();
    let mut slices = Vec::new();
    let mut x = 0.0f32;

    for (visual_order, &logical_index) in visual.iter().enumerate() {
        let c = contributions[logical_index];
        if c.shaped_cluster {
            clusters.push(TypesetClusterAdvanceOut {
                run_index: c.run_index,
                start_char: c.start_char,
                end_char: c.end_char,
                advance: c.advance,
                x_offset: x,
                bidi_level: c.level,
                logical_order: c.logical_order,
            });
        }
        slices.push(TypesetBidiSliceOut {
            run_index: c.run_index,
            start_char: c.start_char,
            end_char: c.end_char,
            advance: c.advance,
            bidi_level: c.level,
            visual_order: visual_order as u32,
            logical_order: c.logical_order,
        });
        if let Some(last) = runs.last_mut().filter(|last| last.run_index == c.run_index) {
            last.start_char = last.start_char.min(c.start_char);
            last.end_char = last.end_char.max(c.end_char);
            last.advance += c.advance;
            last.logical_order = last.logical_order.min(c.logical_order);
        } else {
            runs.push(TypesetRunAdvanceOut {
                run_index: c.run_index,
                start_char: c.start_char,
                end_char: c.end_char,
                advance: c.advance,
                logical_order: c.logical_order,
            });
        }
        x += c.advance;
    }

    (
        (!runs.is_empty()).then_some(runs),
        (!clusters.is_empty()).then_some(clusters),
        (!slices.is_empty()).then_some(slices),
    )
}

/// UTF-16 offset of char index `i` (or the run's total length past the end).
fn utf16_at(t: &PreparedText, i: usize) -> u32 {
    t.chars.get(i).map_or(t.utf16_len, |c| c.utf16_offset)
}

/// Shaped advance sum plus tracking between complete clusters. This is the
/// authoritative path for paint/hit geometry: no gap may land inside a
/// ligature, surrogate pair, or combining sequence.
fn span_width(chars: &[CharAdv], letter_spacing: f32) -> f32 {
    let advance: f32 = chars.iter().map(|c| c.advance).sum();
    if letter_spacing != 0.0 && chars.len() > 1 {
        advance + letter_spacing * (chars.len() - 1) as f32
    } else {
        advance
    }
}

/// TS `findMaxFittingLength`: the longest char-count prefix whose width
/// stays within `max_width`. Char granularity (never inside a surrogate
/// pair); early exit assumes monotonic growth, exact for `ls >= 0`.
fn max_fitting(chars: &[CharAdv], letter_spacing: f32, max_width: f32) -> usize {
    let mut best = 0usize;
    let mut advance = 0.0f32;
    for (k, c) in chars.iter().enumerate() {
        advance += c.advance;
        let cluster_count = k + 1;
        let width = if letter_spacing != 0.0 && cluster_count > 1 {
            advance + letter_spacing * (cluster_count - 1) as f32
        } else {
            advance
        };
        if width <= max_width {
            best = k + 1;
        } else if letter_spacing >= 0.0 {
            break;
        }
    }
    best
}

fn to_flags(compat: &CompatIn) -> wm::CompatFlags {
    wm::CompatFlags {
        no_leading: compat.no_leading,
        do_not_expand_shift_return: compat.do_not_expand_shift_return,
    }
}

/// Map TS `ParagraphSpacing` onto a spacing rule, preserving
/// `calculateTypographyMetrics`'s branch order exactly: `exact`, `atLeast`
/// (both need `line`), then `lineUnit: "multiplier"`, then `lineUnit: "px"`,
/// else single spacing. A `line` value with no recognized `lineUnit`/rule is
/// ignored, like TS. Values are clamped non-negative (security clamp).
fn rule_from_spacing(spacing: Option<&SpacingIn>) -> wm::LineSpacingRule {
    let single = wm::LineSpacingRule::Auto { line_240ths: 240 };
    let Some(sp) = spacing else {
        return single;
    };
    match (sp.line_rule.as_deref(), sp.line, sp.line_unit.as_deref()) {
        (Some("exact"), Some(l), _) => wm::LineSpacingRule::Exact { px: l.max(0.0) },
        (Some("atLeast"), Some(l), _) => wm::LineSpacingRule::AtLeast { px: l.max(0.0) },
        (_, Some(l), Some("multiplier")) => wm::LineSpacingRule::Auto {
            // DOCX multipliers are w:line/240 (240ths round-trip exactly)
            line_240ths: (f64::from(l) * 240.0).round().clamp(0.0, 24_000_000.0) as u32,
        },
        (_, Some(l), Some("px")) => wm::LineSpacingRule::Exact { px: l.max(0.0) },
        _ => single,
    }
}

/// TS: the empty-paragraph single-line floor applies for `auto` (or absent)
/// and `atLeast` rules only.
fn floor_applies(spacing: Option<&SpacingIn>) -> bool {
    matches!(
        spacing.and_then(|sp| sp.line_rule.as_deref()),
        None | Some("auto") | Some("atLeast")
    )
}

#[cfg(test)]
mod tests {
    //! Non-BMP index safety, proven on the production line filler (the code
    //! that emits `headChar`/`tailChar`). The vendored fixture font is
    //! BMP-only (cmap formats 4/6), so a covered emoji cannot flow through
    //! the full pipeline; these tests drive `fill` with a prepared "a😀b"
    //! run — offsets exactly as `prepare` builds them — and assert the
    //! emitted indices count UTF-16 code units and never land inside the
    //! surrogate pair.

    use super::*;
    use crate::font_store::FontStore;

    const FIXTURE: &[u8] = include_bytes!("../../tests/fonts/LiberationSans-Regular.ttf");

    /// "a😀b" with 10px per char: offsets 0/1/3, total UTF-16 length 4.
    fn emoji_run(store: &mut FontStore) -> (PreparedRun, FontId) {
        let id = store.register(FIXTURE.to_vec()).unwrap();
        let t = PreparedText {
            chars: vec![
                CharAdv {
                    utf16_offset: 0,
                    utf16_len: 1,
                    advance: 10.0,
                    level: 0,
                    logical_order: 0,
                    font_size_pt: 12.0,
                    metrics_font: id,
                    baseline_shift_px: 0.0,
                },
                CharAdv {
                    utf16_offset: 1,
                    utf16_len: 2,
                    advance: 10.0,
                    level: 0,
                    logical_order: 1,
                    font_size_pt: 12.0,
                    metrics_font: id,
                    baseline_shift_px: 0.0,
                },
                CharAdv {
                    utf16_offset: 3,
                    utf16_len: 1,
                    advance: 10.0,
                    level: 0,
                    logical_order: 2,
                    font_size_pt: 12.0,
                    metrics_font: id,
                    baseline_shift_px: 0.0,
                },
            ],
            utf16_len: 4,
            breaks: vec![],
            letter_spacing: 0.0,
            font_size_pt: 12.0,
            metrics_font: id,
            baseline_shift_px: 0.0,
        };
        (PreparedRun::Text(t), id)
    }

    fn fill_at(width: f32, prepared: &[PreparedRun]) -> Vec<TypesetRowOut> {
        let compat = CompatIn::default();
        fill(FillParams {
            store: &{
                let mut s = FontStore::new();
                s.register(FIXTURE.to_vec()).unwrap();
                s
            },
            prepared,
            spacing: None,
            body_width: width,
            first_line_width: width,
            default_font_size_pt: 12.0,
            compat: &compat,
            tabs: &[],
            indent_left_px: 0.0,
            first_line_offset_px: 0.0,
            zones: &[],
            paragraph_y_offset: 0.0,
            authoritative_shaping: false,
        })
        .unwrap()
        .lines
    }

    #[test]
    fn hard_break_never_splits_a_surrogate_pair() {
        let mut store = FontStore::new();
        let (run, _) = emoji_run(&mut store);
        // 12px per line: one char each. Cut points must be 1 and 3 — a
        // UTF-16-blind splitter would emit 2 (inside the surrogate pair).
        let lines = fill_at(12.0, std::slice::from_ref(&run));
        let spans: Vec<(u32, u32)> = lines.iter().map(|l| (l.head_char, l.tail_char)).collect();
        assert_eq!(spans, vec![(0, 1), (1, 3), (3, 4)]);
    }

    #[test]
    fn tail_char_counts_utf16_units_not_chars() {
        let mut store = FontStore::new();
        let (run, _) = emoji_run(&mut store);
        // 25px fits two glyphs: the cut after "a😀" is UTF-16 offset 3, not
        // char count 2.
        let lines = fill_at(25.0, std::slice::from_ref(&run));
        let spans: Vec<(u32, u32)> = lines.iter().map(|l| (l.head_char, l.tail_char)).collect();
        assert_eq!(spans, vec![(0, 3), (3, 4)]);
    }

    #[test]
    fn forced_minimum_one_char_takes_the_whole_pair() {
        let mut store = FontStore::new();
        let (run, _) = emoji_run(&mut store);
        // Nothing fits (5px < any glyph): min-1-char lines, and the forced
        // char is the whole emoji (span 1..3), never half of it.
        let lines = fill_at(5.0, std::slice::from_ref(&run));
        let spans: Vec<(u32, u32)> = lines.iter().map(|l| (l.head_char, l.tail_char)).collect();
        assert_eq!(spans, vec![(0, 1), (1, 3), (3, 4)]);
    }
}
