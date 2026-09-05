//! Glyph rasterization for the pre-shaped runs the display list carries.

use std::collections::{HashMap, VecDeque};

use ooxml_text::{FontId, FontStore, PathCmd, shape};
use pptx_render::{PositionedTextLine, PositionedTextRun};
use tiny_skia::{Color, FillRule, Mask, Paint, Path, PathBuilder, Pixmap, Rect, Transform};

use crate::{RenderResources, parse_color};

/// Most glyph outlines one cache retains; the oldest insertion evicts first.
const MAX_CACHED_GLYPH_OUTLINES: usize = 8_192;

/// Extracted glyph outlines keyed by font and glyph id, reusable across slides.
/// Ids are store-local, so the cache binds to the store that fills it and
/// refuses any other; entries past the cap evict oldest-first.
#[derive(Default)]
pub struct GlyphCache {
    entries: HashMap<(u32, u32), CachedGlyph>,
    order: VecDeque<(u32, u32)>,
    store: Option<u64>,
}

impl GlyphCache {
    fn remember(&mut self, key: (u32, u32), glyph: CachedGlyph) {
        if self.entries.insert(key, glyph).is_none() {
            self.order.push_back(key);
        }
        while self.order.len() > MAX_CACHED_GLYPH_OUTLINES {
            let oldest = self
                .order
                .pop_front()
                .expect("the queue tracks every cached entry");
            self.entries.remove(&oldest);
        }
    }
}

#[derive(Debug)]
struct CachedGlyph {
    path: Option<Path>,
    upem: f32,
}

/// Paint one text box's positioned lines. Nothing is shaped here: the layout
/// pass already placed every glyph in slide coordinates.
pub(crate) fn paint_lines(
    pixmap: &mut Pixmap,
    resources: &RenderResources<'_>,
    cache: &mut GlyphCache,
    lines: &[PositionedTextLine],
    transform: Transform,
    clip: Option<&Mask>,
) -> Result<(), String> {
    for line in lines {
        for run in &line.runs {
            paint_run(
                pixmap,
                resources,
                cache,
                run,
                line.baseline,
                transform,
                clip,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_run(
    pixmap: &mut Pixmap,
    resources: &RenderResources<'_>,
    cache: &mut GlyphCache,
    run: &PositionedTextRun,
    baseline: f32,
    transform: Transform,
    clip: Option<&Mask>,
) -> Result<(), String> {
    if !run.font_size_px.is_finite() || run.font_size_px <= 0.0 {
        return Err("run font size must be finite and positive".to_string());
    }
    let mut paint = Paint::default();
    paint.set_color(parse_color(&run.color)?);
    paint.anti_alias = true;

    let font = FontId::from_u32(run.font_id);
    for glyph in &run.glyphs {
        let x = glyph.x + glyph.x_offset;
        if !x.is_finite() || !glyph.y_offset.is_finite() {
            return Err("glyph position must be finite".to_string());
        }
        paint_glyph(
            pixmap,
            resources.fonts,
            cache,
            font,
            glyph.glyph_id,
            run.font_size_px,
            x,
            glyph.y_offset,
            &paint,
            transform,
            clip,
        )?;
    }

    if run.underline {
        paint_underline(pixmap, run, baseline, &paint, transform, clip);
    }
    Ok(())
}

/// The underline geometry the canvas backend draws, so raster and browser
/// underline at the same offset and weight.
fn paint_underline(
    pixmap: &mut Pixmap,
    run: &PositionedTextRun,
    baseline: f32,
    paint: &Paint<'_>,
    transform: Transform,
    clip: Option<&Mask>,
) {
    let Some(rect) = Rect::from_xywh(
        run.x,
        baseline + run.font_size_px * 0.08,
        run.width,
        1.0_f32.max(run.font_size_px * 0.05),
    ) else {
        return;
    };
    pixmap.fill_rect(rect, paint, transform, clip);
}

/// Shape and centre one short label inside `rect`, for the fallback boxes the
/// layout pass emits in place of tables and empty placeholders.
#[allow(clippy::too_many_arguments)]
pub(crate) fn paint_centered_label(
    pixmap: &mut Pixmap,
    fonts: &FontStore,
    cache: &mut GlyphCache,
    font: FontId,
    label: &str,
    size_px: f32,
    color: Color,
    rect: Rect,
    transform: Transform,
    clip: Option<&Mask>,
) -> Result<(), String> {
    let glyphs = shape(fonts, font, label, size_px, &[]).map_err(|error| error.to_string())?;
    let width: f32 = glyphs.iter().map(|glyph| glyph.x_advance).sum();
    let metrics = fonts.metrics(font).map_err(|error| error.to_string())?;
    let upem = f32::from(metrics.units_per_em);
    if upem <= 0.0 {
        return Err("font has no units per em".to_string());
    }
    let ascent = f32::from(metrics.hhea_ascender) / upem * size_px;
    let descent = f32::from(metrics.hhea_descender) / upem * size_px;
    let baseline = rect.y() + rect.height() / 2.0 + (ascent + descent) / 2.0;

    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;

    let mut pen = rect.x() + (rect.width() - width) / 2.0;
    for glyph in &glyphs {
        paint_glyph(
            pixmap,
            fonts,
            cache,
            font,
            glyph.glyph_id,
            size_px,
            pen + glyph.x_offset,
            baseline + glyph.y_offset,
            &paint,
            transform,
            clip,
        )?;
        pen += glyph.x_advance;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_glyph(
    pixmap: &mut Pixmap,
    fonts: &FontStore,
    cache: &mut GlyphCache,
    font: FontId,
    glyph_id: u32,
    size_px: f32,
    x: f32,
    y: f32,
    paint: &Paint<'_>,
    transform: Transform,
    clip: Option<&Mask>,
) -> Result<(), String> {
    let cached = cached_glyph(cache, fonts, font, glyph_id)?;
    let Some(path) = &cached.path else {
        return Ok(());
    };
    let scale = size_px / cached.upem;
    let glyph_transform = Transform::from_row(scale, 0.0, 0.0, -scale, x, y);
    pixmap.fill_path(
        path,
        paint,
        FillRule::Winding,
        transform.pre_concat(glyph_transform),
        clip,
    );
    Ok(())
}

fn cached_glyph<'a>(
    cache: &'a mut GlyphCache,
    fonts: &FontStore,
    font: FontId,
    glyph_id: u32,
) -> Result<&'a CachedGlyph, String> {
    if cache.store.is_some_and(|bound| bound != fonts.id()) {
        return Err("glyph cache is bound to another font store".to_string());
    }
    let key = (font.to_u32(), glyph_id);
    if !cache.entries.contains_key(&key) {
        let id = u16::try_from(glyph_id)
            .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))?;
        let outline = fonts
            .outline_glyph(font, id)
            .map_err(|error| error.to_string())?;
        let path = outline_path(&outline.cmds)?;
        cache.store = Some(fonts.id());
        cache.remember(
            key,
            CachedGlyph {
                path,
                upem: f32::from(outline.upem),
            },
        );
    }
    Ok(cache
        .entries
        .get(&key)
        .expect("a just-inserted or occupied key is present"))
}

fn outline_path(commands: &[PathCmd]) -> Result<Option<Path>, String> {
    if commands.is_empty() {
        return Ok(None);
    }
    let mut builder = PathBuilder::new();
    for command in commands {
        match *command {
            PathCmd::MoveTo { x, y } => builder.move_to(x, y),
            PathCmd::LineTo { x, y } => builder.line_to(x, y),
            PathCmd::QuadTo { cx, cy, x, y } => builder.quad_to(cx, cy, x, y),
            PathCmd::CubicTo {
                c1x,
                c1y,
                c2x,
                c2y,
                x,
                y,
            } => builder.cubic_to(c1x, c1y, c2x, c2y, x, y),
            PathCmd::Close => builder.close(),
        }
    }
    Ok(builder.finish())
}

#[cfg(test)]
mod tests {
    use super::*;

    const CARLITO: &[u8] = include_bytes!("../tests/assets/Carlito-Regular.ttf");

    #[test]
    fn the_cache_evicts_oldest_first_past_its_cap() {
        let mut cache = GlyphCache::default();
        for index in 0..MAX_CACHED_GLYPH_OUTLINES as u32 + 2 {
            cache.remember(
                (index, 0),
                CachedGlyph {
                    path: None,
                    upem: 1_000.0,
                },
            );
        }
        assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
        assert!(!cache.entries.contains_key(&(0, 0)));
        assert!(
            cache
                .entries
                .contains_key(&(MAX_CACHED_GLYPH_OUTLINES as u32 + 1, 0))
        );
    }

    #[test]
    fn a_glyph_cache_is_bound_to_the_store_that_filled_it() {
        let mut first = FontStore::new();
        let font = first.register(CARLITO.to_vec()).expect("register");
        let mut cache = GlyphCache::default();
        let glyph = first.glyph_id(font, 'A').expect("lookup").expect("covered");
        cached_glyph(&mut cache, &first, font, u32::from(glyph)).expect("fill");

        let mut second = FontStore::new();
        let other = second.register(CARLITO.to_vec()).expect("register");
        let error = cached_glyph(&mut cache, &second, other, u32::from(glyph))
            .expect_err("refuse a foreign store");
        assert!(error.contains("bound to another font store"), "{error}");
    }
}
