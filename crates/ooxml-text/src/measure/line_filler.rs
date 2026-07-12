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
use crate::line_metrics as wm;

use super::tabs;

const WRAP_SLACK_PX: f32 = 0.5;
const DEFAULT_SINGLE_LINE_RATIO: f32 = 1.15;

pub(super) struct FillParams<'a> {
    pub store: &'a crate::font_store::FontStore,
    pub prepared: &'a [PreparedRun],
    pub spacing: Option<&'a SpacingIn>,
    /// Content width for every line after the first (indents applied).
    pub body_width: f32,
    /// Content width for the first line (first-line/hanging offset applied).
    pub first_line_width: f32,
    pub default_font_size_pt: f32,
    pub compat: &'a CompatIn,
    /// Custom tab stops (`attrs.tabs`), positions in twips.
    pub tabs: &'a [TabStopIn],
    /// `indent.left` in px — tab stops are content-area-relative, so the
    /// indent is added back when converting line x to grid coordinates.
    pub indent_left_px: f32,
    /// `firstLine − hanging` in px, applied to grid x on the first line only.
    pub first_line_offset_px: f32,
    pub zones: &'a [FloatZoneIn],
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
    max_image_height_px: f32,
    available: f32,
    left_offset: f32,
    /// Float margin from the content right edge.
    right_offset: f32,
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

type AdvanceMetadata = (
    Option<Vec<TypesetRunAdvanceOut>>,
    Option<Vec<TypesetClusterAdvanceOut>>,
    Option<Vec<TypesetBidiSliceOut>>,
);

struct Filler<'a> {
    p: &'a FillParams<'a>,
    rule: wm::LineSpacingRule,
    compat: wm::CompatFlags,
    lines: Vec<TypesetRowOut>,
    cur: LineState,
    cumulative_height: f32,
    pending_float_skip: f32,
}

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

fn estimated_line_height(p: &FillParams) -> f32 {
    pt_to_px(p.default_font_size_pt)
}

pub(super) fn fill(p: FillParams) -> Result<ParagraphExtentOut, MeasureError> {
    let rule = rule_from_spacing(p.spacing);
    let compat = to_flags(p.compat);

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
        line_height = line_height.max(size_px * DEFAULT_SINGLE_LINE_RATIO);
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

    fn fill_tab_run(&mut self, run_index: usize, t: PreparedTab) -> Result<(), MeasureError> {
        let ri = run_index as u32;
        self.update_max_font(t.font_size_pt, t.metrics_font, 0.0);

        let following = self.following_width_after(run_index);
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

        // Clamp overhanging stops while reserving following content.
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

    fn following_width_after(&self, tab_index: usize) -> f32 {
        let mut width = 0.0f32;
        for prun in &self.p.prepared[tab_index + 1..] {
            match prun {
                PreparedRun::Tab(_) | PreparedRun::LineBreak => break,
                PreparedRun::Text(t) => width += span_width(&t.chars, t.letter_spacing),
                PreparedRun::Field(f) => width += f.width,
                PreparedRun::InlineImage(img) => width += img.width,
                PreparedRun::OwnLineImage(img) => width += img.width,
                PreparedRun::SkippedImage { width: w, .. } => width += w,
                PreparedRun::Hidden { .. } => {}
            }
        }
        width
    }

    fn fill_text_run(&mut self, ri: u32, t: &PreparedText) -> Result<(), MeasureError> {
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

        self.cumulative_height += text_line_height;
        Ok(())
    }

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

fn advance_metadata(contributions: &[LineContribution]) -> AdvanceMetadata {
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

fn rule_from_spacing(spacing: Option<&SpacingIn>) -> wm::LineSpacingRule {
    let single = wm::LineSpacingRule::Auto { line_240ths: 240 };
    let Some(sp) = spacing else {
        return single;
    };
    match (sp.line_rule.as_deref(), sp.line, sp.line_unit.as_deref()) {
        (Some("exact"), Some(l), _) => wm::LineSpacingRule::Exact { px: l.max(0.0) },
        (Some("atLeast"), Some(l), _) => wm::LineSpacingRule::AtLeast { px: l.max(0.0) },
        (_, Some(l), Some("multiplier")) => wm::LineSpacingRule::Auto {
            line_240ths: (f64::from(l) * 240.0).round().clamp(0.0, 24_000_000.0) as u32,
        },
        (_, Some(l), Some("px")) => wm::LineSpacingRule::Exact { px: l.max(0.0) },
        _ => single,
    }
}

fn floor_applies(spacing: Option<&SpacingIn>) -> bool {
    matches!(
        spacing.and_then(|sp| sp.line_rule.as_deref()),
        None | Some("auto") | Some("atLeast")
    )
}

#[cfg(test)]
mod tests {
    //! Non-BMP line-boundary coverage.

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
