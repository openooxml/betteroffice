//! Greedy line filling over the per-cluster advance tables built by
//! [`super::prepare`].
//!
//! This module owns every wrap decision and all vertical arithmetic: where a
//! line breaks, how tall its box is, how floats narrow or displace it, and
//! which [`TypesetRowOut`] fields come out. The rules it enforces:
//!
//! - Trailing spaces retain their advance but do not force a word to wrap.
//!   [`WRAP_SLACK_PX`] absorbs floating-point rounding at the line edge.
//! - A word too wide for a whole line is chopped: the current line takes what
//!   fits, then each following line takes the longest cluster prefix that
//!   fits, with a forced minimum of one cluster so filling always terminates.
//!   Cuts land on cluster boundaries, so no ligature, combining sequence or
//!   surrogate pair is ever split.
//! - Line metrics follow the largest font on the line: the first
//!   font-bearing contribution claims it, and only a strictly larger size
//!   displaces it. Ascent, descent and leading are per-contribution maxima,
//!   so even a text run with no characters raises the line box. A line with
//!   no font-bearing run at all falls back to a 0.8/0.2 em ascent/descent
//!   split on a [`DEFAULT_SINGLE_LINE_RATIO`] basis.
//! - The emitted `ascent`/`descent` are the *spacing-ruled* pair, not the raw
//!   content metrics, so `ascent + descent <= lineHeight` always holds: an
//!   `exact`, floored `atLeast` or sub-single box moves the pair rather than
//!   overflowing the box. An image-grown line overrides both and the
//!   identity still holds.
//! - A tall inline image sits on the baseline with text descent below it;
//!   alone on its line it takes the image's own height and nothing more.
//!   Block images retain a descent buffer above and below their footprint.
//! - Float geometry is probed per line at the running Y with a fixed
//!   default-font-size estimate, then re-tested against `fullWidthBlock` bands
//!   once the line closes and its box is known. That running Y advances by each
//!   line's *text* height, so image growth never shifts the next probe; float
//!   skips do, since they move the line itself.
//! - `totalHeight` is Σ (line height + `floatSkipBefore`) plus
//!   `spacing.before` and `spacing.after`.

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

/// Floating-point tolerance at the line edge.
const WRAP_SLACK_PX: f32 = 1e-3;
/// Empty-paragraph line height floor, as a multiple of the font size; applies
/// under the `auto` and `atLeast` rules only.
const WORD_SINGLE_LINE_FLOOR: f32 = 1.15;
/// Single-line basis for a line carrying no font-bearing run.
const DEFAULT_SINGLE_LINE_RATIO: f32 = 1.15;

/// Everything the filler needs that is fixed for the whole paragraph.
pub(super) struct FillParams<'a> {
    pub justify: bool,
    pub store: &'a crate::font_store::FontStore,
    pub prepared: &'a [PreparedRun],
    pub spacing: Option<&'a SpacingIn>,
    /// Content width for every line after the first (indents applied).
    pub body_width: f32,
    /// Content width for the first line (first-line/hanging offset applied).
    pub first_line_width: f32,
    /// Default font size used to seed line metrics.
    pub default_font_size_pt: f32,
    pub compat: &'a CompatIn,
    /// Custom tab stops (`attrs.tabs`), positions in twips.
    pub tabs: &'a [TabStopIn],
    /// `indent.left` in px — tab stops are content-area-relative, so the
    /// indent is added back when converting line x to grid coordinates.
    pub indent_left_px: f32,
    /// `firstLine − hanging` in px, applied to grid x on the first line only.
    pub first_line_offset_px: f32,
    /// Validated float exclusion zones.
    pub zones: &'a [FloatZoneIn],
    /// Paragraph Y in the floating-zone coordinate space.
    pub paragraph_y_offset: f32,
    pub authoritative_shaping: bool,
    /// Grid pitch in px for snap-to-grid (`w:docGrid w:linePitch`), already
    /// gated to an activating grid type AND the paragraph opt-out (`None`
    /// disables snapping). The filler additionally requires an `auto`
    /// spacing rule. Per-line run opt-outs in `run_snaps` can still
    /// disable individual lines.
    pub snap_pitch_px: Option<f32>,
    /// Per prepared run (index-aligned with `prepared`): whether the run
    /// allows grid snapping (`w:snapToGrid`, default on). A line containing
    /// any disallowing run does not snap.
    pub run_snaps: &'a [bool],
}

/// The line currently being filled. Reset wholesale by `start_new_line`.
struct LineState {
    head_run: u32,
    head_char: u32,
    tail_run: u32,
    tail_char: u32,
    width: f32,
    space_width: f32,
    max_font_size_pt: f32,
    max_font: Option<FontId>,
    max_ascent: f32,
    max_descent: f32,
    max_below_baseline: f32,
    /// Tallest inline-image footprint on the line.
    max_image_height_px: f32,
    available: f32,
    /// Float margin from the content left edge.
    left_offset: f32,
    /// Float margin from the content right edge.
    right_offset: f32,
    /// Split strips from centered floating exclusions.
    segment_zones: Option<Vec<FloatSegmentIn>>,
    contributions: Vec<LineContribution>,
}

/// One advance-bearing piece of the current line, recorded in logical order.
/// `logical_order` is `run_index × 1_000_000 + position within the run`, so a
/// single key orders the whole paragraph. `shaped_cluster` distinguishes text
/// clusters (which reach `clusterAdvances`) from atomic runs such as tabs,
/// fields and images (which do not).
#[derive(Debug, Clone, Copy)]
struct LineContribution {
    run_index: u32,
    start_char: u32,
    end_char: u32,
    advance: f32,
    level: u8,
    logical_order: u32,
    shaped_cluster: bool,
    is_space: bool,
    is_fit_space: bool,
}

/// Fill state: the paragraph's finished lines, the line in progress, and the
/// vertical position driving float probes.
struct Filler<'a> {
    p: &'a FillParams<'a>,
    rule: wm::LineSpacingRule,
    compat: wm::CompatFlags,
    lines: Vec<TypesetRowOut>,
    cur: LineState,
    /// Running text-height Y within the paragraph.
    cumulative_height: f32,
    /// Float skip attached to the next finalized line.
    pending_float_skip: f32,
}

/// Hops the running Y past any float leaving less than
/// [`floats::MIN_WRAP_SEGMENT_WIDTH`] of usable room. The skipped pixels are
/// added to both the running Y and the pending `floatSkipBefore`.
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

/// Probe height for zone intersection tests: the default font size in px,
/// never the line's own metrics, which are unknown until the line closes.
fn estimated_line_height(p: &FillParams) -> f32 {
    pt_to_px(p.default_font_size_pt)
}

/// Wraps the prepared runs into lines and totals the paragraph height.
///
/// The first line's float context is resolved before any content is placed:
/// probe at the paragraph top, hop past obstructions, then take the margins
/// at the Y actually reached.
pub(super) fn fill(p: FillParams) -> Result<ParagraphExtentOut, MeasureError> {
    let rule = rule_from_spacing(p.spacing);
    let compat = to_flags(p.compat);

    // Probe initial float margins with the default-font height.
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
            space_width: 0.0,
            max_font_size_pt: p.default_font_size_pt,
            max_font: None,
            max_ascent: 0.0,
            max_descent: 0.0,
            max_below_baseline: 0.0,
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

    // Total height includes line skips and paragraph spacing.
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

/// Measures an empty or whitespace-only paragraph as one zero-width line at
/// the ruled height of `font` at `size_pt`, floored at
/// [`WORD_SINGLE_LINE_FLOOR`] × the font size under every rule but `exact`.
/// When `snap_pitch_px` is set and the rule is `auto`, the content box is
/// first rounded up to a whole number of grid rows, so the rule's multiple
/// scales the quantized pitch (a pinned `exact`/`atLeast` height never
/// snaps).
///
/// `floats` carries the zone list and the paragraph's absolute Y. The line has
/// no width to narrow, so only a `fullWidthBlock` band moves it, and it drops
/// below any band its box reaches.
pub(super) fn empty_paragraph_extent(
    store: &crate::font_store::FontStore,
    font: FontId,
    size_pt: f32,
    spacing: Option<&SpacingIn>,
    compat: &CompatIn,
    snap_pitch_px: Option<f32>,
    floats: (&[FloatZoneIn], f32),
) -> Result<ParagraphExtentOut, MeasureError> {
    let metrics = store
        .metrics(font)
        .map_err(|e| MeasureError::Invalid(e.to_string()))?;
    let size_px = pt_to_px(size_pt);
    let content = wm::single_line_box(metrics, size_px, &to_flags(compat));
    let rule = rule_from_spacing(spacing);
    // Pinned boxes (`exact` fixed, `atLeast` author-floored) never snap;
    // only automatically-determined heights do.
    let auto_rule = matches!(rule, wm::LineSpacingRule::Auto { .. });
    let content = match snap_pitch_px.filter(|_| auto_rule) {
        Some(pitch) => wm::snap_line_box(content, pitch),
        None => content,
    };
    let ruled = wm::apply_spacing_rule(content, &rule);
    let mut line_height = ruled.height();
    if floor_applies(&rule) {
        line_height = line_height.max(size_px * WORD_SINGLE_LINE_FLOOR);
    }

    let (zones, paragraph_y_offset) = floats;
    let skip = floats::clear_full_width_band_y(paragraph_y_offset, line_height, zones)
        - paragraph_y_offset;
    let float_skip_before = (skip > 0.0).then_some(skip);

    let mut total = line_height + float_skip_before.unwrap_or(0.0);
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
            ascent: ruled.ascent,
            descent: ruled.descent,
            line_height,
            left_offset: None,
            right_offset: None,
            segments: None,
            float_skip_before,
            run_advances: None,
            cluster_advances: None,
            bidi_slices: None,
        }],
        total_height: total,
    })
}

impl Filler<'_> {
    /// Places every prepared run in order, then closes the last line. A
    /// paragraph therefore always emits at least one line.
    fn run(&mut self) -> Result<(), MeasureError> {
        for (run_index, prun) in self.p.prepared.iter().enumerate() {
            let ri = run_index as u32;
            match prun {
                PreparedRun::LineBreak => {
                    // A soft return closes before its run and reopens after it.
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
                    // Floating images advance the logical span without width.
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

    /// Place an inline image at its declared extent.
    fn fill_inline_image(&mut self, ri: u32, img: PreparedImage) -> Result<(), MeasureError> {
        if self.cur.width > 0.0 && self.cur.width + img.width > self.cur.available + WRAP_SLACK_PX {
            self.start_new_line(ri, 0)?;
        }
        if img.height > self.cur.max_image_height_px {
            self.cur.max_image_height_px = img.height;
        }
        self.record_atomic(ri, 0, 1, img.width, img.bidi_level);
        self.cur.width += img.width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    /// Gives a block or `topAndBottom` image its own line: finish any line
    /// already carrying content, take the image's *declared* height plus wrap
    /// distances as the whole box (no column fit), add no width, and open a
    /// fresh line. When the image is the paragraph's last run that fresh line
    /// closes empty.
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

    /// Places a premeasured field as one unbreakable unit: it wraps whole to
    /// the next line rather than splitting, but never off an empty line.
    fn fill_field_run(&mut self, ri: u32, f: PreparedField) -> Result<(), MeasureError> {
        if self.cur.width > 0.0 && self.cur.width + f.width > self.cur.available + WRAP_SLACK_PX {
            self.start_new_line(ri, 0)?;
        }
        self.update_max_font(f.font_size_pt, f.metrics_font, f.baseline_shift_px);
        self.record_atomic(ri, 0, 1, f.width, f.bidi_level);
        self.cur.width += f.width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    /// Places a tab against the stop grid, recomputing its width after wrapping.
    fn fill_tab_run(&mut self, run_index: usize, t: PreparedTab) -> Result<(), MeasureError> {
        let ri = run_index as u32;
        let following = self.following_width_after(run_index);
        let mut tab_width = self.tab_width(following);
        if self.cur.width + tab_width > self.cur.available + WRAP_SLACK_PX {
            self.start_new_line(ri, 0)?;
            tab_width = self.tab_width(following);
        }

        self.update_max_font(t.font_size_pt, t.metrics_font, 0.0);
        self.record_atomic(ri, 0, 1, tab_width, t.bidi_level);
        self.cur.width += tab_width;
        self.cur.tail_run = ri;
        self.cur.tail_char = 1;
        Ok(())
    }

    fn tab_width(&self, following: f32) -> f32 {
        let line_x = self.cur.width + self.cur.left_offset;
        let is_first_line = self.lines.is_empty();
        let content_x = self.p.indent_left_px
            + if is_first_line {
                self.p.first_line_offset_px
            } else {
                0.0
            }
            + line_x;
        tabs::calculate_tab_width(
            content_x,
            self.p.tabs,
            tabs::px_to_twips(self.p.indent_left_px),
            following,
            self.cur.available - self.cur.width,
        )
    }

    /// Sums inline widths until the next tab or line break — what `end` and
    /// `center` stops anchor on. Every image kind contributes its declared
    /// width here, floating and own-line images included, even though neither
    /// adds width to the line advance itself.
    fn following_width_after(&self, tab_index: usize) -> f32 {
        let mut width = 0.0f32;
        for prun in &self.p.prepared[tab_index + 1..] {
            match prun {
                PreparedRun::Tab(_) | PreparedRun::LineBreak => break,
                PreparedRun::Text(t) => width += span_width(&t.chars, t.letter_spacing),
                PreparedRun::Field(f) => width += f.width,
                // Every image kind contributes declared width after a tab.
                PreparedRun::InlineImage(img) => width += img.width,
                PreparedRun::OwnLineImage(img) => width += img.width,
                PreparedRun::SkippedImage { width: w, .. } => width += w,
                PreparedRun::Hidden { .. } => {}
            }
        }
        width
    }

    /// Fills a text run word by word, a word being the span up to the next
    /// UAX-14 break opportunity and so including its trailing space. Words
    /// wider than a whole line are chopped at cluster boundaries with a
    /// forced minimum of one cluster per line.
    fn fill_text_run(&mut self, ri: u32, t: &PreparedText) -> Result<(), MeasureError> {
        if t.chars.is_empty() {
            self.update_max_font(t.font_size_pt, t.metrics_font, t.baseline_shift_px);
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
            let fitting_width = visible_span_width(word, t.letter_spacing);

            if fitting_width > self.cur.available + WRAP_SLACK_PX {
                // Overlong unbreakable word: fill the remaining space on the
                // current line, then hard-break with at least one character.
                let mut chunk_start = 0usize;
                while chunk_start < word.len() {
                    let space_left = self.cur.available - self.cur.width + WRAP_SLACK_PX;
                    let remaining = &word[chunk_start..];
                    let mut best = if visible_span_width(remaining, t.letter_spacing) <= space_left
                    {
                        remaining.len()
                    } else {
                        max_fitting(remaining, t.letter_spacing, space_left)
                    };
                    if best == 0 {
                        if self.cur.width > 0.0 {
                            self.start_new_line(ri, utf16_at(t, char_idx + chunk_start))?;
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
                    }
                }
                char_idx = next_break;
                continue;
            }

            if self.cur.width > 0.0
                && fitting_width > 0.0
                && self.cur.width + fitting_width
                    - if self.p.justify {
                        self.cur.space_width * 0.25
                    } else {
                        0.0
                    }
                    > self.cur.available + WRAP_SLACK_PX
            {
                self.start_new_line(ri, utf16_at(t, char_idx))?;
            }

            self.record_text_clusters(ri, word, t.letter_spacing);
            self.cur.width += word_width;
            self.cur.tail_run = ri;
            self.cur.tail_char = utf16_at(t, next_break);
            char_idx = next_break;
        }
        Ok(())
    }

    /// Accumulates each run's extents above and below the shared baseline.
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
            self.cur.max_below_baseline = self
                .cur
                .max_below_baseline
                .max((line.descent + line.leading - baseline_shift_px).max(0.0));
        }
    }

    /// Records an unbreakable run (tab, field, image) as a single
    /// contribution spanning one index unit.
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
            is_space: false,
            is_fit_space: false,
        });
    }

    /// Records a span of shaped clusters, charging letter spacing to every
    /// cluster but the last — so gaps are counted within the span, never
    /// across its trailing edge into the next word.
    fn record_text_clusters(&mut self, run_index: u32, chars: &[CharAdv], spacing: f32) {
        for (index, cluster) in chars.iter().enumerate() {
            if cluster.is_space {
                self.cur.space_width += cluster.advance.max(0.0);
            }
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
                is_space: cluster.is_space,
                is_fit_space: cluster.is_fit_space,
            });
        }
    }

    /// Whether the current line may snap: the paragraph carries an active
    /// grid pitch, the spacing rule leaves the height automatic (`auto` —
    /// a pinned `exact` box is fixed regardless of content and an `atLeast`
    /// floor is author-set, so Word snaps neither), and no contributing
    /// run opts out. Lines with no recorded contributions (e.g. only
    /// hidden runs) defer to the paragraph.
    fn line_may_snap(&self) -> bool {
        if self.p.snap_pitch_px.is_none() {
            return false;
        }
        if !matches!(self.rule, wm::LineSpacingRule::Auto { .. }) {
            return false;
        }
        self.cur.contributions.iter().all(|part| {
            self.p
                .run_snaps
                .get(part.run_index as usize)
                .copied()
                .unwrap_or(true)
        })
    }

    /// Round the content box up to a whole number of grid rows, before the
    /// spacing rule scales it.
    fn snap_content_box(&self, content: wm::LineBox) -> wm::LineBox {
        match (self.p.snap_pitch_px, self.line_may_snap()) {
            (Some(pitch), true) => wm::snap_line_box(content, pitch),
            _ => content,
        }
    }

    /// Snap a final (possibly image-grown) line box. Ascent/descent stay
    /// put; the caller grows only the box.
    fn snap_line_height(&self, height: f32) -> f32 {
        match (self.p.snap_pitch_px, self.line_may_snap()) {
            (Some(pitch), true) => wm::snap_line_height(height, pitch),
            _ => height,
        }
    }

    /// Closes the current line: resolve typography from its largest font,
    /// apply the spacing rule, let any taller image grow the box, attach
    /// float offsets, segments and skip, and push the row. Advances the
    /// running Y by the *text* height only, so image growth never moves the
    /// next line's float probe.
    fn finalize_line(&mut self) -> Result<(), MeasureError> {
        if self.p.justify {
            let end = self
                .cur
                .contributions
                .iter()
                .rposition(|part| !part.is_fit_space)
                .map_or(0, |i| i + 1);
            let parts = &mut self.cur.contributions[..end];
            let visible_width = parts.iter().map(|part| part.advance).sum::<f32>();
            let space_width = parts
                .iter()
                .filter(|part| part.is_space)
                .map(|part| part.advance.max(0.0))
                .sum::<f32>();
            if visible_width > self.cur.available && space_width > 0.0 {
                let compression = (visible_width - self.cur.available).min(space_width * 0.25);
                for part in parts.iter_mut().filter(|part| part.is_space) {
                    part.advance -= part.advance.max(0.0) / space_width * compression;
                }
                self.cur.width -= compression;
            }
        }
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
                leading: self.cur.max_below_baseline - self.cur.max_descent,
            },
            // Fontless lines use a 0.8/0.2 em split.
            None => wm::LineBox {
                ascent: size_px * 0.8,
                descent: size_px * 0.2,
                leading: size_px * (DEFAULT_SINGLE_LINE_RATIO - 1.0),
            },
        };
        let content = self.snap_content_box(content);
        let ruled = wm::apply_spacing_rule(content, &self.rule);
        let mut ascent = ruled.ascent;
        let mut descent = ruled.descent;
        let text_line_height = ruled.height();
        let mut line_height = text_line_height;

        // An image dictates the whole box, so it buffers from the content
        // descent and reports that as the row descent; the spacing rule no
        // longer describes this line.
        if self.cur.max_image_height_px > line_height {
            let image_h = self.cur.max_image_height_px;
            let buffer = content.descent;
            descent = buffer;
            if self.cur.head_run == self.cur.tail_run
                && matches!(
                    self.p.prepared.get(self.cur.head_run as usize),
                    Some(PreparedRun::OwnLineImage(_))
                )
            {
                line_height = image_h + buffer * 2.0;
                ascent = image_h + buffer;
            } else if self.cur.max_font.is_none() {
                // Word's box for an inline image alone on its line is exactly
                // the image: the paragraph mark buys no descent under it.
                descent = 0.0;
                line_height = image_h;
                ascent = image_h;
            } else {
                line_height = image_h + buffer;
                ascent = image_h;
            }
            // The grid snaps the final box of an `auto`-ruled line,
            // whatever grew it (`line_may_snap` still gates pinned rules
            // out). Ascent/descent stay put so the extra lands below the
            // descent.
            line_height = self.snap_line_height(line_height);
        }

        self.clear_full_width_bands(line_height);

        // Float fields are omitted when unset.
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
            descent,
            line_height,
            left_offset: (self.cur.left_offset > 0.0).then_some(self.cur.left_offset),
            right_offset: (self.cur.right_offset > 0.0).then_some(self.cur.right_offset),
            segments,
            float_skip_before,
            run_advances,
            cluster_advances,
            bidi_slices,
        });

        // Float probes advance by text height, excluding image growth.
        self.cumulative_height += text_line_height;
        Ok(())
    }

    /// Drops the closed line below any band its box reaches, taking the
    /// margins it lands in. Narrower room declines: the fill would overflow.
    fn clear_full_width_bands(&mut self, line_height: f32) {
        if self.p.zones.is_empty() {
            return;
        }
        let top = self.p.paragraph_y_offset + self.cumulative_height;
        let skip = floats::clear_full_width_band_y(top, line_height, self.p.zones) - top;
        if skip <= 0.0 {
            return;
        }
        let full = if self.lines.is_empty() {
            self.p.first_line_width
        } else {
            self.p.body_width
        };
        let margins = floats::floating_margins(
            self.cumulative_height + skip,
            line_height,
            self.p.zones,
            self.p.paragraph_y_offset,
        );
        let available = floats::available_width(&margins, full).max(1.0);
        if available + WRAP_SLACK_PX < self.cur.available {
            return;
        }
        self.cumulative_height += skip;
        self.pending_float_skip += skip;
        self.cur.available = available;
        self.cur.left_offset = margins.left;
        self.cur.right_offset = margins.right;
        self.cur.segment_zones = margins.segments;
    }

    /// Splits the just-closed line across the zone's strips. One strip — or a
    /// line that fits the first strip within the wrap slack — yields a single
    /// segment covering the whole line. A two-way split needs a line made of
    /// exactly one text run; the cut is the longest cluster prefix fitting the
    /// first strip. Anything else (multiple runs, a non-text run, a degenerate
    /// cut) returns `None` and the line carries no segments.
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

        // Character offsets are absolute within the run.
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

    /// Closes the current line and opens a fresh one at `(run, char_utf16)`,
    /// hopped past and narrowed by whatever float zones sit at the new Y.
    /// Font tracking resets, so the new line's metrics start from scratch.
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
            space_width: 0.0,
            max_font_size_pt: self.p.default_font_size_pt,
            max_font: None,
            max_ascent: 0.0,
            max_descent: 0.0,
            max_below_baseline: 0.0,
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

/// Reorders a line's contributions visually and derives the three advance
/// tables. Cluster and slice x offsets accumulate in visual order; run
/// advances merge contributions that are adjacent *visually* and share a run
/// index, so a run cut by a direction change appears once per visual piece.
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

fn visible_span_width(chars: &[CharAdv], letter_spacing: f32) -> f32 {
    let end = chars
        .iter()
        .rposition(|cluster| !cluster.is_fit_space)
        .map_or(0, |i| i + 1);
    span_width(&chars[..end], letter_spacing)
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

/// Longest cluster prefix whose width stays within `max_width`, in clusters.
/// The early exit assumes width grows monotonically, which holds for
/// non-negative letter spacing; negative spacing scans the whole span.
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
        gdi_line_metrics: {
            #[cfg(test)]
            {
                compat.gdi_line_metrics
            }
            #[cfg(not(test))]
            {
                false
            }
        },
        typo_line_spacing: {
            #[cfg(test)]
            {
                compat.typo_line_spacing
            }
            #[cfg(not(test))]
            {
                false
            }
        },
    }
}

/// Maps `w:spacing` onto a line rule, in precedence order: `exact`,
/// `atLeast` (both needing a `line` value), then `lineUnit` `multiplier`,
/// then `px`, else single spacing. A `line` with no recognized unit or rule
/// is ignored, and values are clamped non-negative.
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

/// Tests whether the empty-paragraph line-height floor applies. It reads the
/// *resolved* rule, not the raw `lineRule` string, so `lineUnit: "px"` — which
/// resolves to `exact` — is not floored above the box it asked for.
fn floor_applies(rule: &wm::LineSpacingRule) -> bool {
    !matches!(rule, wm::LineSpacingRule::Exact { .. })
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
                    is_space: false,
                    is_fit_space: false,
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
                    is_space: false,
                    is_fit_space: false,
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
                    is_space: false,
                    is_fit_space: false,
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
        let run_snaps = vec![true; prepared.len()];
        fill(FillParams {
            justify: false,
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
            snap_pitch_px: None,
            run_snaps: &run_snaps,
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
