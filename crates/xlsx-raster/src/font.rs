//! embedded font stack: shapes (rustybuzz) and rasterizes single-line text runs
//! with a bundled carlito regular (metric-compatible with calibri). deterministic.

use std::sync::LazyLock;

use rustybuzz::ttf_parser::{self, OutlineBuilder};
use tiny_skia::{Color, FillRule, Paint, Path, PathBuilder, Pixmap, Rect, Transform};

use unicode_segmentation::UnicodeSegmentation;
use xlsx_render::{Align, NumberFormatLayout, Rect as DlRect, TextValign};

// synthetic bold nudges a second glyph pass right; synthetic italic shears.
const BOLD_OFFSET_PX: f32 = 0.35;
const ITALIC_SHEAR: f32 = 0.21;

const FONT_BYTES: &[u8] = include_bytes!("../assets/Carlito-Regular.ttf");

static FACE: LazyLock<rustybuzz::Face<'static>> = LazyLock::new(|| {
    rustybuzz::Face::from_slice(FONT_BYTES, 0).expect("embedded carlito is a valid font")
});

const PX_PER_PT: f32 = 96.0 / 72.0;
pub const STACKED_LINE_HEIGHT_RATIO: f32 = 1.2;
const STACKED_ASCENT_RATIO: f32 = 0.8;

pub fn stacked_line_height(font_size_pt: f32) -> f32 {
    font_size_pt * PX_PER_PT * STACKED_LINE_HEIGHT_RATIO
}

fn rotated_available_length(cell: &DlRect, deg: f32, font_size_pt: f32) -> f32 {
    xlsx_render::rotated_available_length(cell.w, cell.h, deg, font_size_pt)
}

/// advance width in px of `text` at `font_size_pt` with the embedded font.
pub fn measure_text(text: &str, font_size_pt: f32) -> f32 {
    let face = &*FACE;
    let scale = font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let glyphs = shape(face, text);
    let total: i32 = glyphs.glyph_positions().iter().map(|p| p.x_advance).sum();
    total as f32 * scale
}

/// a single-line text run; `x`/`y` are the align anchor and alphabetic baseline.
pub struct TextRun<'a> {
    pub x: f32,
    pub y: f32,
    pub text: &'a str,
    pub font_size_pt: f32,
    pub color: Color,
    pub align: Align,
    pub clip: &'a DlRect,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub highlight: Option<Color>,
    pub dashed_underline: bool,
}

/// paint one shaped, single-line run, clipped strictly to the run's clip rect.
pub fn paint_text(pixmap: &mut Pixmap, run: &TextRun) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 {
        return;
    }

    let face = &*FACE;
    let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let glyphs = shape(face, run.text);

    let total: i32 = glyphs.glyph_positions().iter().map(|p| p.x_advance).sum();
    let pen = match run.align {
        Align::Left => run.x,
        Align::Center => run.x - total as f32 * scale / 2.0,
        Align::Right => run.x - total as f32 * scale,
    };

    let mut paint = Paint::default();
    paint.set_color(run.color);
    paint.anti_alias = true;

    // tiny-skia mask clipping costs O(pixmap) per run; runs inside the clip
    // paint directly, overflowing runs paint into a clip-sized scratch and blit.
    let ascent = face.ascender() as f32 * scale;
    let descent = face.descender() as f32 * scale;
    let run_width = total as f32 * scale;
    let fully_inside = pen >= run.clip.x
        && pen + run_width <= run.clip.x + run.clip.w
        && run.y - ascent >= run.clip.y
        && run.y - descent <= run.clip.y + run.clip.h;

    if fully_inside {
        paint_run(
            pixmap, face, &glyphs, pen, run, scale, &paint, run_width, 0.0, 0.0,
        );
        return;
    }

    let ox = run.clip.x.floor();
    let oy = run.clip.y.floor();
    let sw = ((run.clip.x + run.clip.w).ceil() - ox).max(1.0) as u32;
    let sh = ((run.clip.y + run.clip.h).ceil() - oy).max(1.0) as u32;
    let Some(mut scratch) = Pixmap::new(sw, sh) else {
        return;
    };
    paint_run(
        &mut scratch,
        face,
        &glyphs,
        pen,
        run,
        scale,
        &paint,
        run_width,
        ox,
        oy,
    );
    pixmap.draw_pixmap(
        ox as i32,
        oy as i32,
        scratch.as_ref(),
        &tiny_skia::PixmapPaint::default(),
        Transform::identity(),
        None,
    );
}

#[allow(clippy::too_many_arguments)]
fn paint_run(
    target: &mut Pixmap,
    face: &rustybuzz::Face<'_>,
    glyphs: &rustybuzz::GlyphBuffer,
    pen: f32,
    run: &TextRun,
    scale: f32,
    paint: &Paint,
    run_width: f32,
    ox: f32,
    oy: f32,
) {
    paint_highlight(target, face, run, pen, run_width, scale, ox, oy);
    paint_glyphs(
        target, face, glyphs, pen, run.y, scale, paint, run.bold, run.italic, ox, oy,
    );
    paint_decorations(target, face, run, pen, run_width, scale, paint, ox, oy);
}

#[allow(clippy::too_many_arguments)]
fn paint_highlight(
    target: &mut Pixmap,
    face: &rustybuzz::Face<'_>,
    run: &TextRun,
    pen: f32,
    run_width: f32,
    scale: f32,
    ox: f32,
    oy: f32,
) {
    let Some(color) = run.highlight else {
        return;
    };
    let ascent = face.ascender() as f32 * scale;
    let descent = -face.descender() as f32 * scale;
    let Some(rect) = Rect::from_xywh(
        pen - ox - 2.0,
        run.y - oy - ascent - 1.0,
        run_width + 4.0,
        ascent + descent + 2.0,
    ) else {
        return;
    };
    let mut paint = Paint::default();
    paint.set_color(color);
    target.fill_rect(rect, &paint, Transform::identity(), None);
}

// (ox, oy) translates into a clip-sized scratch when the run overflows its clip.
#[allow(clippy::too_many_arguments)]
fn paint_glyphs(
    pixmap: &mut Pixmap,
    face: &rustybuzz::Face<'_>,
    glyphs: &rustybuzz::GlyphBuffer,
    start_pen: f32,
    baseline: f32,
    scale: f32,
    paint: &Paint,
    bold: bool,
    italic: bool,
    ox: f32,
    oy: f32,
) {
    let shear = if italic { ITALIC_SHEAR * scale } else { 0.0 };
    let mut pen = start_pen;
    for (info, pos) in glyphs.glyph_infos().iter().zip(glyphs.glyph_positions()) {
        if let Some(path) = glyph_path(face, info.glyph_id as u16) {
            // outlines are y-up in font units; -scale flips into y-down pixel space.
            let tx = pen + pos.x_offset as f32 * scale - ox;
            let ty = baseline - pos.y_offset as f32 * scale - oy;
            let transform = Transform::from_row(scale, 0.0, shear, -scale, tx, ty);
            pixmap.fill_path(&path, paint, FillRule::Winding, transform, None);
            if bold {
                let embolden = transform.post_translate(BOLD_OFFSET_PX, 0.0);
                pixmap.fill_path(&path, paint, FillRule::Winding, embolden, None);
            }
        }
        pen += pos.x_advance as f32 * scale;
    }
}

#[allow(clippy::too_many_arguments)]
fn paint_decorations(
    target: &mut Pixmap,
    face: &rustybuzz::Face<'_>,
    run: &TextRun,
    pen: f32,
    run_width: f32,
    scale: f32,
    paint: &Paint,
    ox: f32,
    oy: f32,
) {
    if run_width <= 0.0 {
        return;
    }
    let em = face.units_per_em() as f32;
    let mut bar = |center_up: f32, thickness: f32| {
        let h = (thickness * scale).max(0.5);
        let cy = run.y - center_up * scale - oy;
        if let Some(rect) = Rect::from_xywh(pen - ox, cy - h / 2.0, run_width, h) {
            target.fill_rect(rect, paint, Transform::identity(), None);
        }
    };
    if run.underline {
        let m = face.underline_metrics();
        let pos = m.map(|m| m.position as f32).unwrap_or(-0.1 * em);
        let thick = m.map(|m| m.thickness as f32).unwrap_or(0.05 * em);
        bar(pos, thick);
    }
    if run.strike {
        let m = face.strikeout_metrics();
        let pos = m.map(|m| m.position as f32).unwrap_or(0.26 * em);
        let thick = m.map(|m| m.thickness as f32).unwrap_or(0.05 * em);
        bar(pos, thick);
    }
    if run.dashed_underline {
        let m = face.underline_metrics();
        let pos = m.map(|m| m.position as f32).unwrap_or(-0.1 * em);
        let thick = m.map(|m| m.thickness as f32).unwrap_or(0.05 * em);
        let h = (thick * scale).max(0.5);
        let y = run.y - pos * scale - oy - h / 2.0;
        let dash = 3.0_f32.max(h * 2.0);
        let gap = 2.0_f32.max(h);
        let mut x = pen - ox;
        let end = x + run_width;
        while x < end {
            if let Some(rect) = Rect::from_xywh(x, y, dash.min(end - x), h) {
                target.fill_rect(rect, paint, Transform::identity(), None);
            }
            x += dash + gap;
        }
    }
}

pub fn paint_rotated_text(
    pixmap: &mut Pixmap,
    run: &TextRun,
    deg: f32,
    cell: &DlRect,
    valign: TextValign,
    number_format: Option<&NumberFormatLayout>,
) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 || cell.w <= 0.0 || cell.h <= 0.0 {
        return;
    }
    let face = &*FACE;
    let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let (segments, run_width) = match number_format {
        Some(layout) => {
            let mut effective = layout.clone();
            effective.left = 0.0;
            effective.width = rotated_available_length(cell, deg, run.font_size_pt);
            let mut raw: Vec<(String, f32)> = Vec::new();
            xlsx_render::paint_number_format(
                run.text,
                &effective,
                run.align,
                |literal| measure_text(literal, run.font_size_pt),
                |literal, x| raw.push((literal.to_owned(), x)),
            );
            if raw.is_empty() {
                (
                    vec![(run.text.to_owned(), 0.0)],
                    measure_text(run.text, run.font_size_pt),
                )
            } else {
                let min_x = raw.iter().map(|(_, x)| *x).fold(f32::INFINITY, f32::min);
                let max_x = raw
                    .iter()
                    .map(|(text, x)| x + measure_text(text, run.font_size_pt))
                    .fold(f32::NEG_INFINITY, f32::max);
                let segments = raw
                    .into_iter()
                    .map(|(text, x)| (text, x - min_x))
                    .collect::<Vec<_>>();
                (segments, (max_x - min_x).max(0.0))
            }
        }
        None => (
            vec![(run.text.to_owned(), 0.0)],
            measure_text(run.text, run.font_size_pt),
        ),
    };
    let ascent = face.ascender() as f32 * scale;
    let descent = -face.descender() as f32 * scale;
    let line_h = ascent + descent;
    let a = -deg.to_radians();
    let (sin, cos) = a.sin_cos();
    let bbox_w = (run_width * cos).abs() + (line_h * sin).abs();
    let bbox_h = (run_width * sin).abs() + (line_h * cos).abs();
    let bbox_left = match run.align {
        Align::Left => cell.x,
        Align::Center => cell.x + (cell.w - bbox_w) / 2.0,
        Align::Right => cell.x + cell.w - bbox_w,
    };
    let bbox_top = match valign {
        TextValign::Top => cell.y,
        TextValign::Center => cell.y + (cell.h - bbox_h) / 2.0,
        TextValign::Bottom => cell.y + cell.h - bbox_h,
    };
    let cx = bbox_left + bbox_w / 2.0;
    let cy = bbox_top + bbox_h / 2.0;
    let ox = run.clip.x.floor();
    let oy = run.clip.y.floor();
    let sw = ((run.clip.x + run.clip.w).ceil() - ox).max(1.0) as u32;
    let sh = ((run.clip.y + run.clip.h).ceil() - oy).max(1.0) as u32;
    let Some(mut scratch) = Pixmap::new(sw, sh) else {
        return;
    };
    let mut paint = Paint::default();
    paint.set_color(run.color);
    paint.anti_alias = true;
    let center_y_local = (-ascent + descent) / 2.0;
    if let Some(color) = run.highlight {
        let mut hp = Paint::default();
        hp.set_color(color);
        if let Some(rect) = Rect::from_xywh(
            -run_width / 2.0 - 2.0,
            -ascent - center_y_local - 1.0,
            run_width + 4.0,
            line_h + 2.0,
        ) {
            let transform = Transform::from_row(cos, sin, -sin, cos, cx - ox, cy - oy);
            scratch.fill_rect(rect, &hp, transform, None);
        }
    }
    let shear = if run.italic {
        ITALIC_SHEAR * scale
    } else {
        0.0
    };
    for (segment, segment_offset) in &segments {
        let glyphs = shape(face, segment);
        let mut pen = 0.0_f32;
        for (info, pos) in glyphs.glyph_infos().iter().zip(glyphs.glyph_positions()) {
            if let Some(path) = glyph_path(face, info.glyph_id as u16) {
                let tx0 = segment_offset + pen + pos.x_offset as f32 * scale - run_width / 2.0;
                let ty0 = -pos.y_offset as f32 * scale - center_y_local;
                let e = cos * tx0 - sin * ty0 + cx - ox;
                let f = sin * tx0 + cos * ty0 + cy - oy;
                let transform = Transform::from_row(
                    cos * scale,
                    sin * scale,
                    cos * shear + sin * scale,
                    sin * shear - cos * scale,
                    e,
                    f,
                );
                scratch.fill_path(&path, &paint, FillRule::Winding, transform, None);
                if run.bold {
                    let bold = Transform::from_row(
                        cos * scale,
                        sin * scale,
                        cos * shear + sin * scale,
                        sin * shear - cos * scale,
                        e + cos * BOLD_OFFSET_PX,
                        f + sin * BOLD_OFFSET_PX,
                    );
                    scratch.fill_path(&path, &paint, FillRule::Winding, bold, None);
                }
            }
            pen += pos.x_advance as f32 * scale;
        }
    }
    if run.underline || run.strike || run.dashed_underline {
        let em = face.units_per_em() as f32;
        let mut bar = |center_up: f32, thickness: f32| {
            let h = (thickness * scale).max(0.5);
            let y_local = -center_up * scale - center_y_local;
            if let Some(rect) = Rect::from_xywh(-run_width / 2.0, y_local - h / 2.0, run_width, h) {
                let transform = Transform::from_row(cos, sin, -sin, cos, cx - ox, cy - oy);
                scratch.fill_rect(rect, &paint, transform, None);
            }
        };
        if run.underline || run.dashed_underline {
            let m = face.underline_metrics();
            bar(
                m.map(|m| m.position as f32).unwrap_or(-0.1 * em),
                m.map(|m| m.thickness as f32).unwrap_or(0.05 * em),
            );
        }
        if run.strike {
            let m = face.strikeout_metrics();
            bar(
                m.map(|m| m.position as f32).unwrap_or(0.26 * em),
                m.map(|m| m.thickness as f32).unwrap_or(0.05 * em),
            );
        }
    }
    pixmap.draw_pixmap(
        ox as i32,
        oy as i32,
        scratch.as_ref(),
        &tiny_skia::PixmapPaint::default(),
        Transform::identity(),
        None,
    );
}

pub fn paint_stacked_text(pixmap: &mut Pixmap, run: &TextRun, cell: &DlRect, valign: TextValign) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 {
        return;
    }
    let line_h = stacked_line_height(run.font_size_pt);
    let ascent = run.font_size_pt * PX_PER_PT * STACKED_ASCENT_RATIO;
    let count = run.text.graphemes(true).count();
    let total_h = count as f32 * line_h;
    let top = match valign {
        TextValign::Top => cell.y,
        TextValign::Center => cell.y + (cell.h - total_h) / 2.0,
        TextValign::Bottom => cell.y + cell.h - total_h,
    };
    let clip_top = run.clip.y;
    let clip_bottom = run.clip.y + run.clip.h;
    for (i, cluster) in run.text.graphemes(true).enumerate() {
        let line_top = top + i as f32 * line_h;
        if line_top + line_h <= clip_top || line_top >= clip_bottom {
            continue;
        }
        let w = measure_text(cluster, run.font_size_pt);
        let x = match run.align {
            Align::Left => cell.x,
            Align::Center => cell.x + (cell.w - w) / 2.0,
            Align::Right => cell.x + cell.w - w,
        };
        if x + w <= run.clip.x || x >= run.clip.x + run.clip.w {
            continue;
        }
        let sub = TextRun {
            x,
            y: line_top + ascent,
            text: cluster,
            font_size_pt: run.font_size_pt,
            color: run.color,
            align: Align::Left,
            clip: run.clip,
            bold: run.bold,
            italic: run.italic,
            underline: run.underline,
            strike: run.strike,
            highlight: run.highlight,
            dashed_underline: run.dashed_underline,
        };
        paint_text(pixmap, &sub);
    }
}

fn shape<'a>(face: &rustybuzz::Face<'a>, text: &str) -> rustybuzz::GlyphBuffer {
    let mut buffer = rustybuzz::UnicodeBuffer::new();
    buffer.push_str(text);
    rustybuzz::shape(face, &[], buffer)
}

fn glyph_path(face: &ttf_parser::Face, glyph_id: u16) -> Option<Path> {
    let mut builder = PathCollector {
        pb: PathBuilder::new(),
    };
    face.outline_glyph(ttf_parser::GlyphId(glyph_id), &mut builder)?;
    builder.pb.finish()
}

struct PathCollector {
    pb: PathBuilder,
}

impl OutlineBuilder for PathCollector {
    fn move_to(&mut self, x: f32, y: f32) {
        self.pb.move_to(x, y);
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.pb.line_to(x, y);
    }

    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.pb.quad_to(x1, y1, x, y);
    }

    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.pb.cubic_to(x1, y1, x2, y2, x, y);
    }

    fn close(&mut self) {
        self.pb.close();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn measure_is_positive_and_scales() {
        let small = measure_text("Hello", 11.0);
        let big = measure_text("Hello", 22.0);
        assert!(small > 0.0);
        assert!((big - small * 2.0).abs() < 0.01);
    }

    #[test]
    fn empty_string_has_zero_width() {
        assert_eq!(measure_text("", 11.0), 0.0);
    }

    #[test]
    fn wider_text_measures_wider() {
        assert!(measure_text("mmmm", 11.0) > measure_text("ii", 11.0));
    }
}
