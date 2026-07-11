//! embedded font stack: shapes (rustybuzz) and rasterizes single-line text runs
//! with a bundled carlito regular (metric-compatible with calibri). deterministic.

use std::sync::LazyLock;

use rustybuzz::ttf_parser::{self, OutlineBuilder};
use tiny_skia::{Color, FillRule, Paint, Path, PathBuilder, Pixmap, Rect, Transform};

use xlsx_render::{Align, Rect as DlRect};

// synthetic bold nudges a second glyph pass right; synthetic italic shears.
const BOLD_OFFSET_PX: f32 = 0.35;
const ITALIC_SHEAR: f32 = 0.21;

const FONT_BYTES: &[u8] = include_bytes!("../assets/Carlito-Regular.ttf");

static FACE: LazyLock<rustybuzz::Face<'static>> = LazyLock::new(|| {
    rustybuzz::Face::from_slice(FONT_BYTES, 0).expect("embedded carlito is a valid font")
});

const PX_PER_PT: f32 = 96.0 / 72.0;

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
    paint_glyphs(
        target, face, glyphs, pen, run.y, scale, paint, run.bold, run.italic, ox, oy,
    );
    paint_decorations(target, face, run, pen, run_width, scale, paint, ox, oy);
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
