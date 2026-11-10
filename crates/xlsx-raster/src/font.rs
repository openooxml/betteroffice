//! embedded font stack: shapes and rasterizes single-line text runs with a
//! bundled carlito regular (metric-compatible with calibri). deterministic.

use std::cell::RefCell;

use ooxml_raster::{CachedGlyph, GlyphCache};
use ooxml_text::{FontId, FontStore, ShapedGlyph};
use tiny_skia::{Color, FillRule, Paint, Pixmap, Rect, Transform};

use unicode_segmentation::UnicodeSegmentation;
use xlsx_render::{Align, NumberFormatLayout, Rect as DlRect, TextValign};

use crate::fmath;

// synthetic bold nudges a second glyph pass right; synthetic italic shears.
const BOLD_OFFSET_PX: f32 = 0.35;
const ITALIC_SHEAR: f32 = 0.21;

const FONT_BYTES: &[u8] = include_bytes!("../assets/Carlito-Regular.ttf");

/// A face's metrics as rustybuzz reads them, taken once when the face is
/// registered: units per em, ascender, descender, and the underline and
/// strikeout position and thickness when the font has them.
#[derive(Clone, Copy, Debug)]
pub(crate) struct FaceInfo {
    upem: i32,
    ascender: i16,
    descender: i16,
    underline: Option<(i16, i16)>,
    strikeout: Option<(i16, i16)>,
}

impl FaceInfo {
    pub(crate) fn read(bytes: &[u8]) -> Option<Self> {
        let face = rustybuzz::Face::from_slice(bytes, 0)?;
        Some(Self {
            upem: face.units_per_em(),
            ascender: face.ascender(),
            descender: face.descender(),
            underline: face
                .underline_metrics()
                .map(|metrics| (metrics.position, metrics.thickness)),
            strikeout: face
                .strikeout_metrics()
                .map(|metrics| (metrics.position, metrics.thickness)),
        })
    }
}

/// A face a run draws with: its id in a font store, its metrics, and the
/// glyph cache its outlines go through, which is bound to that store.
#[derive(Clone, Copy)]
pub(crate) struct RunFace<'a> {
    pub(crate) store: &'a FontStore,
    pub(crate) id: FontId,
    pub(crate) info: FaceInfo,
    pub(crate) cache: &'a RefCell<GlyphCache>,
}

impl RunFace<'_> {
    fn units_per_em(&self) -> i32 {
        self.info.upem
    }

    fn ascender(&self) -> i16 {
        self.info.ascender
    }

    fn descender(&self) -> i16 {
        self.info.descender
    }

    fn underline_metrics(&self) -> Option<(i16, i16)> {
        self.info.underline
    }

    fn strikeout_metrics(&self) -> Option<(i16, i16)> {
        self.info.strikeout
    }
}

thread_local! {
    static EMBEDDED: EmbeddedFont = EmbeddedFont::new();
}

/// The bundled carlito in a store of its own, one per thread.
struct EmbeddedFont {
    store: FontStore,
    id: FontId,
    info: FaceInfo,
    cache: RefCell<GlyphCache>,
}

impl EmbeddedFont {
    fn new() -> Self {
        let mut store = FontStore::new();
        let id = store
            .register_slice(FONT_BYTES)
            .expect("embedded carlito is a valid font");
        Self {
            store,
            id,
            info: FaceInfo::read(FONT_BYTES).expect("embedded carlito is a valid font"),
            cache: RefCell::new(GlyphCache::default()),
        }
    }
}

/// Runs `paint` with the embedded carlito.
pub(crate) fn with_embedded<R>(paint: impl FnOnce(RunFace<'_>) -> R) -> R {
    EMBEDDED.with(|font| {
        paint(RunFace {
            store: &font.store,
            id: font.id,
            info: font.info,
            cache: &font.cache,
        })
    })
}

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
    with_embedded(|face| {
        let scale = font_size_pt * PX_PER_PT / face.units_per_em() as f32;
        let glyphs = shape(face, text);
        let total = total_advance(&glyphs);
        total as f32 * scale
    })
}
#[cfg(not(feature = "fonts"))]
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
#[cfg(not(feature = "fonts"))]
/// paint one shaped, single-line run, clipped strictly to the run's clip rect.
pub fn paint_text(pixmap: &mut Pixmap, run: &TextRun) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 {
        return;
    }
    with_embedded(|face| paint_text_with_face(pixmap, run, face));
}

#[cfg(not(feature = "fonts"))]
fn paint_text_with_face(pixmap: &mut Pixmap, run: &TextRun, face: RunFace<'_>) {
    let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let glyphs = shape(face, run.text);

    let total = total_advance(&glyphs);
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
    face: RunFace<'_>,
    glyphs: &[ShapedGlyph],
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
    face: RunFace<'_>,
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
    face: RunFace<'_>,
    glyphs: &[ShapedGlyph],
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
    let mut cache = face.cache.borrow_mut();
    for pos in glyphs {
        if let Some(path) = glyph_path(face, &mut cache, pos.glyph_id) {
            // outlines are y-up in font units; -scale flips into y-down pixel space.
            let tx = pen + pos.x_offset * scale - ox;
            let ty = baseline - pos.y_offset * scale - oy;
            let transform = Transform::from_row(scale, 0.0, shear, -scale, tx, ty);
            pixmap.fill_path(path, paint, FillRule::Winding, transform, None);
            if bold {
                let embolden = transform.post_translate(BOLD_OFFSET_PX, 0.0);
                pixmap.fill_path(path, paint, FillRule::Winding, embolden, None);
            }
        }
        pen += pos.x_advance * scale;
    }
}

#[allow(clippy::too_many_arguments)]
fn paint_decorations(
    target: &mut Pixmap,
    face: RunFace<'_>,
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
        let pos = m.map(|m| m.0 as f32).unwrap_or(-0.1 * em);
        let thick = m.map(|m| m.1 as f32).unwrap_or(0.05 * em);
        bar(pos, thick);
    }
    if run.strike {
        let m = face.strikeout_metrics();
        let pos = m.map(|m| m.0 as f32).unwrap_or(0.26 * em);
        let thick = m.map(|m| m.1 as f32).unwrap_or(0.05 * em);
        bar(pos, thick);
    }
    if run.dashed_underline {
        let m = face.underline_metrics();
        let pos = m.map(|m| m.0 as f32).unwrap_or(-0.1 * em);
        let thick = m.map(|m| m.1 as f32).unwrap_or(0.05 * em);
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
#[cfg(not(feature = "fonts"))]
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
    with_embedded(|face| {
        paint_rotated_text_with_face(pixmap, run, deg, cell, valign, number_format, face)
    });
}

#[cfg(not(feature = "fonts"))]
fn paint_rotated_text_with_face(
    pixmap: &mut Pixmap,
    run: &TextRun,
    deg: f32,
    cell: &DlRect,
    valign: TextValign,
    number_format: Option<&NumberFormatLayout>,
    face: RunFace<'_>,
) {
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
    let (sin, cos) = fmath::sin_cos(a);
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
    let mut cache = face.cache.borrow_mut();
    for (segment, segment_offset) in &segments {
        let glyphs = shape(face, segment);
        let mut pen = 0.0_f32;
        for pos in &glyphs {
            if let Some(path) = glyph_path(face, &mut cache, pos.glyph_id) {
                let tx0 = segment_offset + pen + pos.x_offset * scale - run_width / 2.0;
                let ty0 = -pos.y_offset * scale - center_y_local;
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
                scratch.fill_path(path, &paint, FillRule::Winding, transform, None);
                if run.bold {
                    let bold = Transform::from_row(
                        cos * scale,
                        sin * scale,
                        cos * shear + sin * scale,
                        sin * shear - cos * scale,
                        e + cos * BOLD_OFFSET_PX,
                        f + sin * BOLD_OFFSET_PX,
                    );
                    scratch.fill_path(path, &paint, FillRule::Winding, bold, None);
                }
            }
            pen += pos.x_advance * scale;
        }
    }
    drop(cache);
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
                m.map(|m| m.0 as f32).unwrap_or(-0.1 * em),
                m.map(|m| m.1 as f32).unwrap_or(0.05 * em),
            );
        }
        if run.strike {
            let m = face.strikeout_metrics();
            bar(
                m.map(|m| m.0 as f32).unwrap_or(0.26 * em),
                m.map(|m| m.1 as f32).unwrap_or(0.05 * em),
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
#[cfg(not(feature = "fonts"))]
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

/// `text` shaped at the face's units per em, so advances and offsets come
/// back as exact font units.
pub(crate) fn shape(face: RunFace<'_>, text: &str) -> Vec<ShapedGlyph> {
    ooxml_text::shape(face.store, face.id, text, face.units_per_em() as f32, &[])
        .unwrap_or_default()
}

/// The summed advances of `glyphs`, in font units, added as integers.
pub(crate) fn total_advance(glyphs: &[ShapedGlyph]) -> i32 {
    glyphs.iter().map(|glyph| glyph.x_advance as i32).sum()
}

/// The outline of `glyph_id`, built once per face and kept in the cache.
pub(crate) fn glyph_path<'a>(
    face: RunFace<'_>,
    cache: &'a mut GlyphCache,
    glyph_id: u32,
) -> Option<&'a tiny_skia::Path> {
    let CachedGlyph { path, .. } = cache
        .lookup(
            face.store,
            face.id,
            glyph_id,
            || u16::try_from(glyph_id).map_err(|error| error.to_string()),
            |cmds| Ok(ooxml_raster::outline_path(cmds)),
        )
        .ok()?;
    path.as_ref()
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

    const TEXTS: [&str; 9] = [
        "",
        "Hello, world!",
        "   ",
        "AV To ffi office",
        "e\u{301} A\u{30a}",
        "مرحبا بالعالم",
        "אבג 123",
        "中文",
        "1,234.50 $ (12%)",
    ];
    const SIZES: [f32; 8] = [-11.0, -0.0, 0.0, 1.0, 8.5, 11.0, 22.0, 72.0];

    fn rustybuzz_width(face: &rustybuzz::Face<'_>, text: &str, size: f32) -> f32 {
        let mut buffer = rustybuzz::UnicodeBuffer::new();
        buffer.push_str(text);
        let glyphs = rustybuzz::shape(face, &[], buffer);
        let total: i32 = glyphs.glyph_positions().iter().map(|p| p.x_advance).sum();
        total as f32 * (size * PX_PER_PT / face.units_per_em() as f32)
    }

    /// The shared stack measures exactly what rustybuzz did, bit for bit,
    /// including a run long enough to overflow an f32 sum of font units.
    #[test]
    fn measure_text_matches_rustybuzz() {
        let face = rustybuzz::Face::from_slice(FONT_BYTES, 0).unwrap();
        let long = "a".repeat(32_767);
        for text in TEXTS.iter().copied().chain(std::iter::once(long.as_str())) {
            for size in SIZES {
                assert_eq!(
                    measure_text(text, size).to_bits(),
                    rustybuzz_width(&face, text, size).to_bits(),
                    "{text:?} at {size}pt"
                );
            }
        }
    }

    /// The metrics taken at registration are the font's own.
    #[test]
    fn face_info_matches_the_font() {
        let face = rustybuzz::Face::from_slice(FONT_BYTES, 0).unwrap();
        with_embedded(|font| {
            assert_eq!(font.units_per_em(), face.units_per_em());
            assert_eq!(font.ascender(), face.ascender());
            assert_eq!(font.descender(), face.descender());
            assert_eq!(
                font.underline_metrics(),
                face.underline_metrics().map(|m| (m.position, m.thickness))
            );
            assert_eq!(
                font.strikeout_metrics(),
                face.strikeout_metrics().map(|m| (m.position, m.thickness))
            );
        });
    }

    /// A glyph's outline is built once and then served from the cache.
    #[test]
    fn glyph_outlines_are_cached() {
        with_embedded(|font| {
            let glyphs = shape(font, "aa");
            let mut cache = font.cache.borrow_mut();
            let first = glyph_path(font, &mut cache, glyphs[0].glyph_id).cloned();
            let again = glyph_path(font, &mut cache, glyphs[1].glyph_id).cloned();
            assert!(first.is_some());
            assert_eq!(first, again);
        });
    }

    /// Through a registered face, the bundled-font path measures what
    /// rustybuzz did on the same bytes.
    #[cfg(feature = "fonts")]
    #[test]
    fn measure_with_registered_faces_matches_rustybuzz() {
        let mut fonts = crate::RasterFonts::default();
        fonts
            .register_font("Carlito", false, false, FONT_BYTES.to_vec(), false, false)
            .unwrap();
        let face = rustybuzz::Face::from_slice(FONT_BYTES, 0).unwrap();
        let (faces, stack) = fonts.chain_with_stack("Carlito", false, false);
        for text in [
            "Hello, world!",
            "AV To ffi office",
            "e\u{301} A\u{30a}",
            "1,234.50 $ (12%)",
        ] {
            let mut buffer = rustybuzz::UnicodeBuffer::new();
            buffer.push_str(text);
            let glyphs = rustybuzz::shape(&face, &[], buffer);
            let total: i32 = glyphs.glyph_positions().iter().map(|p| p.x_advance).sum();
            for size in SIZES {
                let expected = total as f32 * size * PX_PER_PT / face.units_per_em() as f32;
                assert_eq!(
                    measure_text_with_fonts(faces, stack, text, size).to_bits(),
                    expected.to_bits(),
                    "{text:?} at {size}pt"
                );
            }
        }
    }
}

#[cfg(feature = "fonts")]
#[path = "font_native.rs"]
mod native;
#[cfg(feature = "fonts")]
pub use native::FontStack;
#[cfg(feature = "fonts")]
pub(crate) use native::{FontFace, measure_text_with_fonts};
#[cfg(feature = "fonts")]
pub use native::{TextRun, paint_rotated_text, paint_stacked_text, paint_text};
