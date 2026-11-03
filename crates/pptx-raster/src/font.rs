//! Glyph rasterization for the pre-shaped runs the display list carries.

use ooxml_text::{FontId, FontStore, PathCmd, shape};
use pptx_render::{PositionedTextLine, PositionedTextRun};
use tiny_skia::{Color, Mask, Paint, Path, Pixmap, Rect, Transform};

use crate::{RenderResources, parse_color};

/// Store-bound glyph outlines reusable across slides.
#[derive(Default)]
pub struct GlyphCache {
    inner: ooxml_raster::GlyphCache,
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
        paint_underline(
            pixmap,
            run,
            baseline - run.baseline_offset_px,
            &paint,
            transform,
            clip,
        );
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
    ooxml_raster::fill_glyph(pixmap, path, scale, x, y, transform, paint, clip);
    Ok(())
}

fn cached_glyph<'a>(
    cache: &'a mut GlyphCache,
    fonts: &FontStore,
    font: FontId,
    glyph_id: u32,
) -> Result<&'a ooxml_raster::CachedGlyph, String> {
    cache.inner.lookup(
        fonts,
        font,
        glyph_id,
        || {
            u16::try_from(glyph_id)
                .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))
        },
        outline_path,
    )
}

fn outline_path(commands: &[PathCmd]) -> Result<Option<Path>, String> {
    Ok(ooxml_raster::outline_path(commands))
}

#[cfg(test)]
mod tests {
    use super::*;

    const CARLITO: &[u8] = include_bytes!("../tests/assets/Carlito-Regular.ttf");

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

#[cfg(test)]
mod shared_helper_tests {
    use super::*;

    #[test]
    fn outline_keeps_empty_and_undrawable_results() {
        assert!(outline_path(&[]).unwrap().is_none());
        assert!(
            outline_path(&[PathCmd::MoveTo { x: 0.0, y: 0.0 }])
                .unwrap()
                .is_none()
        );
    }
}
