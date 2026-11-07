//! Glyph rasterization for the pre-shaped runs the display list carries.

use ooxml_text::{FontId, FontStore, PathCmd, ShapeFeature, ShapedGlyph, shape};
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
    let features: &[ShapeFeature] = if run.kerning { &[] } else { &[UNKERNED] };
    let fitted = fitted_pieces(run, baseline - run.baseline_offset_px, |text| {
        shape(resources.fonts, font, text, run.font_size_px, features).ok()
    });
    let laid_out = || {
        run.glyphs.iter().map(|glyph| FittedGlyph {
            glyph_id: glyph.glyph_id,
            x: glyph.x + glyph.x_offset,
            y: glyph.y_offset,
            stretch: 1.0,
        })
    };
    let glyphs: Box<dyn Iterator<Item = FittedGlyph>> = match fitted {
        Some(fitted) => Box::new(fitted.into_iter()),
        None => Box::new(laid_out()),
    };
    for glyph in glyphs {
        if !glyph.x.is_finite() || !glyph.y.is_finite() {
            return Err("glyph position must be finite".to_string());
        }
        paint_glyph(
            pixmap,
            resources.fonts,
            cache,
            font,
            glyph.glyph_id,
            run.font_size_px,
            glyph.x,
            glyph.y,
            &paint,
            transform.pre_concat(Transform::from_row(
                glyph.stretch,
                0.0,
                0.0,
                1.0,
                glyph.x - glyph.x * glyph.stretch,
                0.0,
            )),
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
    let strike_y = baseline - run.baseline_offset_px - run.font_size_px * 0.3;
    let thickness = 1.0_f32.max(run.font_size_px * 0.05);
    let offsets: &[f32] = match run.strike.as_deref() {
        Some("sngStrike") => &[0.0],
        Some("dblStrike") => &[-thickness, thickness],
        _ => &[],
    };
    for offset in offsets {
        if let Some(rect) = Rect::from_xywh(run.x, strike_y + offset, run.width, thickness) {
            pixmap.fill_rect(rect, &paint, transform, clip);
        }
    }
    Ok(())
}

/// The feature a run laid out without kerning is measured and painted with, as the canvas sets `fontKerning`.
const UNKERNED: ShapeFeature = ShapeFeature {
    tag: *b"kern",
    value: 0,
};
const FIT_MIN: f64 = 0.75;
const FIT_MAX: f64 = 1.45;

/// Whitespace as the canvas backend's `\s` reads it: ECMAScript white space and line terminators.
fn is_space(ch: char) -> bool {
    matches!(ch as u32,
        0x09..=0x0d | 0x20 | 0xa0 | 0x1680 | 0x2000..=0x200a | 0x2028 | 0x2029 | 0x202f | 0x205f | 0x3000 | 0xfeff)
}

/// How far a cluster may sit from the pen before it departs: 0.0001 px, or two single-precision steps at that x.
fn departure_tolerance(expected: f32) -> f32 {
    0.0001_f32.max(expected.abs() * 2f32.powi(-21))
}

fn is_cjk(ch: char) -> bool {
    matches!(ch as u32,
        0x3000..=0x3098 | 0x309b..=0x30ff | 0x3400..=0x4dbf | 0x4e00..=0x9fff | 0xac00..=0xd7af |
        0xf900..=0xfaff | 0xff00..=0xffef | 0x20000..=0x2ffff)
}

/// A glyph to paint: its id, pen x, baseline y and horizontal stretch about its origin.
#[derive(Debug, Clone, Copy, PartialEq)]
struct FittedGlyph {
    glyph_id: u32,
    x: f32,
    y: f32,
    stretch: f32,
}

/// One piece of a fitted run: its text and glyph ranges, where it starts and where the next one does.
struct Piece {
    text_start: usize,
    text_end: usize,
    glyph_start: usize,
    glyph_end: usize,
    x: f32,
    end_x: f64,
}

/// The pieces of a run laid out on another family's widths, split where a glyph departs from the pen of its own
/// advances, exactly as the canvas backend splits it; `None` for a run that is not fitted. Every decision reads the
/// run's single-precision values, which the canvas backend reads too.
fn pieces(run: &PositionedTextRun) -> Option<(Vec<u16>, Vec<Piece>)> {
    use unicode_properties::{GeneralCategoryGroup, UnicodeGeneralCategory};
    let glyphs = &run.glyphs;
    if glyphs.is_empty() || run.letter_spacing_px != 0.0 {
        return None;
    }
    if glyphs
        .windows(2)
        .any(|pair| pair[1].cluster < pair[0].cluster)
    {
        return None;
    }
    let departs = |index: usize| {
        if index > 0 && glyphs[index].cluster == glyphs[index - 1].cluster {
            return false;
        }
        let expected = index.checked_sub(1).map_or(run.x, |previous| {
            glyphs[previous].x + glyphs[previous].advance
        });
        (glyphs[index].x - expected).abs() > departure_tolerance(expected)
    };
    if !(1..glyphs.len()).any(departs) {
        return None;
    }
    let mut lefts: Vec<(u32, f32)> = Vec::new();
    for glyph in glyphs {
        match lefts.last_mut() {
            Some((cluster, left)) if *cluster == glyph.cluster => *left = left.min(glyph.x),
            _ => lefts.push((glyph.cluster, glyph.x)),
        }
    }
    if let Some(&(_, first)) = lefts.first()
        && lefts[1..]
            .iter()
            .find(|(_, left)| (f64::from(*left) - f64::from(first)).abs() > 0.01)
            .is_some_and(|(_, left)| *left < first)
    {
        return None;
    }
    let units: Vec<u16> = run.text.encode_utf16().collect();
    let length = units.len();
    let char_at = |offset: usize| -> Option<char> {
        char::decode_utf16(units.get(offset..)?.iter().copied())
            .next()?
            .ok()
    };
    let char_before = |offset: usize| -> Option<char> {
        let start = if offset >= 2
            && (0xdc00..=0xdfff).contains(&units[offset - 1])
            && (0xd800..=0xdbff).contains(&units[offset - 2])
        {
            offset - 2
        } else {
            offset.checked_sub(1)?
        };
        char_at(start)
    };
    let continuation = |offset: usize| {
        char_at(offset).is_some_and(|ch| {
            ch.general_category_group() == GeneralCategoryGroup::Mark
                || ch == '\u{200c}'
                || ch == '\u{200d}'
        }) || char_before(offset) == Some('\u{200d}')
    };
    let boundary = |offset: usize, departs: bool| {
        char_before(offset).is_some_and(is_space)
            || char_at(offset).is_some_and(is_cjk)
            || char_before(offset).is_some_and(is_cjk)
            || (departs && char_at(offset).is_some_and(is_space))
    };
    let offset = |index: usize| (glyphs[index].cluster.saturating_sub(run.start)) as usize;
    let fitted = (1..glyphs.len()).any(|index| {
        let at = offset(index);
        glyphs[index - 1].advance > 0.0
            && at > 0
            && at < length
            && departs(index)
            && !continuation(at)
            && !boundary(at, true)
    });
    if !fitted {
        return None;
    }
    let mut pieces = Vec::new();
    let mut text_start = 0;
    let mut glyph_start = 0;
    let mut x = run.x;
    for (index, glyph) in glyphs.iter().enumerate() {
        let at = offset(index);
        if at > text_start && at < length && departs(index) {
            pieces.push(Piece {
                text_start,
                text_end: at,
                glyph_start,
                glyph_end: index,
                x,
                end_x: f64::from(glyph.x),
            });
            text_start = at;
            glyph_start = index;
            x = glyph.x;
        }
    }
    pieces.push(Piece {
        text_start,
        text_end: length,
        glyph_start,
        glyph_end: glyphs.len(),
        x,
        end_x: f64::from(run.x) + f64::from(run.width),
    });
    Some((units, pieces))
}

/// The stretch of a piece measured `own` wide, as the canvas backend scales its pieces: the span to the next piece
/// over the measured width, left at 1 for CJK, tab or degenerate pieces and outside 0.75 to 1.45.
fn piece_stretch(piece: &Piece, text: &[u16], own: f64) -> f64 {
    let cjk = char::decode_utf16(text.iter().copied())
        .filter_map(Result::ok)
        .any(is_cjk);
    let tab = text.contains(&u16::from(b'\t'));
    let width = piece.end_x - f64::from(piece.x);
    let fit = if cjk || tab || own <= 0.0 || width <= 0.0 {
        1.0
    } else {
        width / own
    };
    if (fit - 1.0).abs() < 0.001 || !(FIT_MIN..=FIT_MAX).contains(&fit) {
        1.0
    } else {
        fit
    }
}

/// The glyphs to paint for a run laid out on another family's widths: each piece drawn from its own shaping,
/// stretched to its span, as the canvas backend draws it, so what is measured is what is painted. `None` for a run
/// whose glyphs follow their own advances, which paints as laid out.
fn fitted_pieces(
    run: &PositionedTextRun,
    baseline: f32,
    shape_piece: impl Fn(&str) -> Option<Vec<ShapedGlyph>>,
) -> Option<Vec<FittedGlyph>> {
    let (units, pieces) = pieces(run)?;
    let glyphs = &run.glyphs;
    let mut placed = Vec::with_capacity(glyphs.len());
    for piece in &pieces {
        let text = &units[piece.text_start.min(units.len())..piece.text_end.min(units.len())];
        let laid_out = &glyphs[piece.glyph_start..piece.glyph_end];
        if text.contains(&u16::from(b'\t')) {
            placed.extend(laid_out.iter().map(|glyph| FittedGlyph {
                glyph_id: glyph.glyph_id,
                x: glyph.x + glyph.x_offset,
                y: glyph.y_offset,
                stretch: 1.0,
            }));
            continue;
        }
        let shaped = shape_piece(&String::from_utf16_lossy(text));
        let own: f32 = match &shaped {
            Some(shaped) => shaped.iter().map(|glyph| glyph.x_advance).sum(),
            None => laid_out.iter().map(|glyph| glyph.advance).sum(),
        };
        let stretch = piece_stretch(piece, text, f64::from(own)) as f32;
        let mut pen = piece.x;
        match shaped {
            Some(shaped) => {
                for glyph in shaped {
                    placed.push(FittedGlyph {
                        glyph_id: glyph.glyph_id,
                        x: pen + glyph.x_offset * stretch,
                        y: baseline + glyph.y_offset,
                        stretch,
                    });
                    pen += glyph.x_advance * stretch;
                }
            }
            None => {
                for glyph in laid_out {
                    placed.push(FittedGlyph {
                        glyph_id: glyph.glyph_id,
                        x: pen + glyph.x_offset * stretch,
                        y: glyph.y_offset,
                        stretch,
                    });
                    pen += glyph.advance * stretch;
                }
            }
        }
    }
    Some(placed)
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

    fn run(text: &str, glyphs: &[(u32, f32, f32)], width: f32) -> PositionedTextRun {
        PositionedTextRun {
            text: text.into(),
            start: 0,
            end: text.encode_utf16().count() as u32,
            x: 40.0,
            width,
            font_id: 0,
            font_family: "Corbel".into(),
            font_size_px: 20.0,
            bold: false,
            italic: false,
            underline: false,
            strike: None,
            color: "#000000".into(),
            letter_spacing_px: 0.0,
            baseline_offset_px: 0.0,
            kerning: true,
            glyphs: glyphs
                .iter()
                .enumerate()
                .map(
                    |(index, &(cluster, x, advance))| pptx_render::PositionedGlyph {
                        glyph_id: index as u32 + 1,
                        cluster,
                        x,
                        advance,
                        x_offset: 0.0,
                        y_offset: 0.0,
                    },
                )
                .collect(),
        }
    }

    fn per_unit(advance: f32) -> impl Fn(&str) -> Option<Vec<ShapedGlyph>> {
        move |text: &str| {
            Some(
                text.encode_utf16()
                    .enumerate()
                    .map(|(index, _)| ShapedGlyph {
                        glyph_id: 100 + index as u32,
                        cluster: index as u32,
                        x_advance: advance,
                        x_offset: 0.0,
                        y_offset: 0.0,
                    })
                    .collect(),
            )
        }
    }

    #[test]
    fn named_width_pieces_are_fitted_like_the_canvas() {
        let placed = fitted_pieces(
            &run(
                "ab cd",
                &[
                    (0, 40.0, 8.0),
                    (1, 49.0, 8.0),
                    (2, 58.0, 8.0),
                    (3, 66.0, 8.0),
                    (4, 75.0, 8.0),
                ],
                240.0,
            ),
            0.0,
            per_unit(8.0),
        )
        .unwrap();
        let expected = [
            (40.0, 9.0 / 8.0),
            (49.0, 9.0 / 8.0),
            (58.0, 17.0 / 16.0),
            (58.0 + 8.0 * 17.0 / 16.0, 17.0 / 16.0),
            (75.0, 1.0),
        ];
        for (glyph, (want_x, want_stretch)) in placed.iter().zip(expected) {
            assert!(
                (glyph.x - want_x).abs() < 1e-4 && (glyph.stretch - want_stretch).abs() < 1e-4,
                "{placed:?}"
            );
        }
    }

    #[test]
    fn a_run_on_its_own_advances_paints_as_laid_out() {
        assert!(
            fitted_pieces(
                &run(
                    "ab cd",
                    &[
                        (0, 40.0, 8.0),
                        (1, 48.0, 8.0),
                        (2, 56.0, 8.0),
                        (3, 64.0, 8.0),
                        (4, 72.0, 8.0)
                    ],
                    40.0
                ),
                0.0,
                per_unit(8.0)
            )
            .is_none()
        );
    }

    #[test]
    fn a_piece_is_measured_on_its_own_like_the_canvas_measures_it() {
        let placed = fitted_pieces(
            &run("ab", &[(0, 40.0, 6.0), (1, 49.0, 6.0)], 18.0),
            0.0,
            per_unit(8.0),
        )
        .unwrap();
        assert!((placed[0].stretch - 9.0 / 8.0).abs() < 1e-4, "{placed:?}");
    }

    #[test]
    fn the_stretch_limit_is_compared_on_single_precision_values() {
        let span = 5.8_f32;
        let mut at_limit = run("##", &[(0, 0.0, 3.0), (1, span, 3.0)], 2.0 * span);
        at_limit.x = 0.0;
        let placed = fitted_pieces(&at_limit, 0.0, per_unit(4.0)).unwrap();
        let fit = f64::from(span) / 4.0;
        assert!(
            fit > FIT_MAX,
            "5.8 in single precision is a little over 5.8"
        );
        assert_eq!(placed[0].stretch, 1.0, "{placed:?}");
    }

    #[test]
    fn a_departure_after_a_mark_is_not_fitted() {
        let mut marked = run(
            "cafe\u{301}s more",
            &[
                (0, 42.0, 10.0),
                (1, 52.0, 10.0),
                (2, 62.0, 10.0),
                (3, 72.0, 10.0),
                (4, 78.0, 0.0),
                (5, 82.0, 10.0),
                (6, 92.0, 5.0),
                (7, 97.0, 10.0),
                (8, 107.0, 10.0),
                (9, 117.0, 10.0),
                (10, 127.0, 10.0),
            ],
            95.0,
        );
        marked.x = 42.0;
        assert!(fitted_pieces(&marked, 0.0, per_unit(8.0)).is_none());
    }

    #[test]
    fn a_position_its_decimal_rounds_is_not_a_departure() {
        let mut rounded = run(
            "aTo",
            &[
                (0, 2_040.0, 9.0),
                (1, 2_050.0, 10.379_883),
                (2, 2_060.38, 10.0),
            ],
            30.379_883,
        );
        rounded.x = 2040.0;
        assert_eq!(
            rounded.glyphs[2].x,
            rounded.glyphs[1].x + rounded.glyphs[1].advance
        );
        let placed = fitted_pieces(&rounded, 0.0, per_unit(10.0)).unwrap();
        assert_eq!(placed.len(), 3, "{placed:?}");
        let to = placed[1].stretch;
        assert!(
            (placed[2].x - (placed[1].x + 10.0 * to)).abs() < 1e-3,
            "T and o stay one piece: {placed:?}"
        );
    }

    #[test]
    fn a_piece_paints_the_shaping_it_was_measured_with() {
        let kerned = |text: &str| {
            Some(
                text.encode_utf16()
                    .enumerate()
                    .map(|(index, _)| ShapedGlyph {
                        glyph_id: 900 + index as u32,
                        cluster: index as u32,
                        x_advance: 7.0,
                        x_offset: 0.0,
                        y_offset: -2.0,
                    })
                    .collect(),
            )
        };
        let placed = fitted_pieces(
            &run(
                "aTT",
                &[(0, 40.0, 9.0), (1, 50.0, 10.0), (2, 60.0, 10.0)],
                30.0,
            ),
            68.0,
            kerned,
        )
        .unwrap();
        assert_eq!(placed.len(), 3, "{placed:?}");
        assert_eq!(placed[1].glyph_id, 900, "{placed:?}");
        assert_eq!(placed[1].y, 66.0, "{placed:?}");
        let fit = placed[1].stretch;
        assert!((fit - 20.0 / 14.0).abs() < 1e-4 || fit == 1.0, "{placed:?}");
        assert!(
            (placed[2].x - (50.0 + 7.0 * fit)).abs() < 1e-4,
            "{placed:?}"
        );
    }

    #[test]
    fn an_unkerned_run_shapes_its_pieces_without_kerning() {
        let mut fonts = FontStore::new();
        let font = fonts.register(CARLITO.to_vec()).expect("register");
        let kerned = shape(&fonts, font, "AV", 40.0, &[]).unwrap();
        let unkerned = shape(&fonts, font, "AV", 40.0, &[UNKERNED]).unwrap();
        let width =
            |glyphs: &[ShapedGlyph]| glyphs.iter().map(|glyph| glyph.x_advance).sum::<f32>();
        assert!(width(&kerned) < width(&unkerned), "the test face kerns AV");
        let mut laid_out = run(
            "xAV",
            &[(0, 40.0, 20.0), (1, 62.0, 30.0), (2, 92.0, 30.0)],
            82.0,
        );
        laid_out.kerning = false;
        let features: &[ShapeFeature] = if laid_out.kerning { &[] } else { &[UNKERNED] };
        let placed = fitted_pieces(&laid_out, 0.0, |text| {
            shape(&fonts, font, text, 40.0, features).ok()
        })
        .unwrap();
        let piece = (40.0 + 82.0 - 62.0) / width(&unkerned);
        let fit = if (FIT_MIN..=FIT_MAX).contains(&piece) {
            piece
        } else {
            1.0
        };
        assert!(
            (placed[2].x - (62.0 + unkerned[0].x_advance * fit)).abs() < 1e-3,
            "{placed:?}"
        );
    }

    #[test]
    fn a_tab_piece_keeps_its_laid_out_glyphs() {
        let mut tabbed = run("\tab", &[(1, 70.0, 8.0), (2, 81.0, 8.0)], 49.0);
        tabbed.x = 40.0;
        let placed = fitted_pieces(&tabbed, 0.0, per_unit(8.0)).unwrap();
        assert_eq!(placed.len(), 2, "the tab piece paints nothing: {placed:?}");
        assert_eq!(placed[0].x, 70.0, "{placed:?}");
    }

    /// Seeded random runs at large x and odd fractions, with tabs, marks and CJK, and the pieces this engine fits
    /// them in. The canvas backend's tests read the same file and must reach the same decisions.
    #[test]
    fn piece_decisions_match_the_shared_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/piece-decisions.json"
        );
        let mut seed = 0x2545_f491_4f6c_dd1d_u64;
        let mut next = move || {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            (seed >> 11) as f64 / (1u64 << 53) as f64
        };
        let alphabet = [
            'a', 'T', 'o', 'V', 'W', '.', ' ', '\t', '\u{301}', '\u{6f22}', '\u{85}', '\u{a0}',
        ];
        let mut cases = Vec::new();
        for _ in 0..400 {
            let length = 2 + (next() * 11.0) as usize;
            let text: String = (0..length)
                .map(|_| alphabet[(next() * alphabet.len() as f64) as usize % alphabet.len()])
                .collect();
            let start = (next() * 4.0) as u32;
            let x = (next() * 6000.0) as f32 + (next() as f32);
            let mut pen = x;
            let mut glyph_list = Vec::new();
            for (index, ch) in text.chars().enumerate() {
                let roll = next();
                let shift = if roll < 0.55 {
                    0.0
                } else if roll < 0.75 {
                    (next() * 3.0) as f32
                } else if roll < 0.9 {
                    (next() * 0.0004) as f32
                } else {
                    -(next() * 1.5) as f32
                };
                let advance = if ch == '\u{301}' {
                    0.0
                } else {
                    3.0 + (next() * 12.0) as f32 + next() as f32 / 7.0
                };
                let glyph_x = pen + shift;
                glyph_list.push((start + index as u32, glyph_x, advance));
                pen = glyph_x + advance;
                if ch != '\u{301}' && next() < 0.15 {
                    let inner = (next() * 4.0) as f32;
                    let inner_x = pen + (next() * 0.0003) as f32;
                    glyph_list.push((start + index as u32, inner_x, inner));
                    pen = inner_x + inner;
                }
            }
            let order = next();
            if order < 0.08 {
                let clusters: Vec<u32> = glyph_list.iter().rev().map(|glyph| glyph.0).collect();
                for (glyph, cluster) in glyph_list.iter_mut().zip(clusters) {
                    glyph.0 = cluster;
                }
            } else if order < 0.14 {
                let positions: Vec<f32> = glyph_list.iter().rev().map(|glyph| glyph.1).collect();
                for (glyph, x) in glyph_list.iter_mut().zip(positions) {
                    glyph.1 = x;
                }
            } else if order < 0.2 && glyph_list.len() > 2 {
                let at = 1 + (next() * (glyph_list.len() - 2) as f64) as usize;
                let (left, right) = (glyph_list[at].0, glyph_list[at + 1].0);
                glyph_list[at].0 = right;
                glyph_list[at + 1].0 = left;
            }
            let mut case = run(&text, &glyph_list, pen - x + (next() * 5.0) as f32);
            case.start = start;
            case.end = start + text.encode_utf16().count() as u32;
            case.x = x;
            let decisions = pieces(&case).map(|(units, pieces)| {
                pieces
                    .iter()
                    .map(|piece| {
                        let text = &units[piece.text_start..piece.text_end];
                        (
                            piece.text_start,
                            piece_stretch(piece, text, text.len() as f64 * 8.0),
                        )
                    })
                    .collect::<Vec<_>>()
            });
            cases.push(serde_json::json!({ "run": case, "decisions": decisions }));
        }
        let fixture = serde_json::to_string_pretty(
            &serde_json::json!({ "measure": "8 px per UTF-16 unit", "cases": cases }),
        )
        .unwrap()
            + "\n";
        if std::env::var_os("BO_WRITE_PIECE_FIXTURE").is_some() {
            std::fs::create_dir_all(concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures"))
                .unwrap();
            std::fs::write(path, &fixture).unwrap();
        }
        let committed = std::fs::read_to_string(path)
            .expect("run with BO_WRITE_PIECE_FIXTURE=1 to write the fixture");
        assert_eq!(
            committed, fixture,
            "piece decisions changed: rerun with BO_WRITE_PIECE_FIXTURE=1 and check the canvas test"
        );
        let fitted = cases
            .iter()
            .filter(|case| !case["decisions"].is_null())
            .count();
        assert!(
            fitted > 50 && fitted < 380,
            "the fixture mixes fitted and plain runs: {fitted}"
        );
        let out_of_order = cases
            .iter()
            .filter(|case| {
                let clusters: Vec<u64> = case["run"]["glyphs"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|glyph| glyph["cluster"].as_u64().unwrap())
                    .collect();
                clusters.windows(2).any(|pair| pair[1] < pair[0])
            })
            .collect::<Vec<_>>();
        assert!(
            out_of_order.len() > 20,
            "the fixture has runs whose clusters run backwards"
        );
        assert!(out_of_order.iter().all(|case| case["decisions"].is_null()));
    }

    #[test]
    fn a_glyph_inside_a_cluster_never_departs() {
        let mut syllable = run(
            "ab",
            &[
                (0, 1_039.4, 8.39),
                (0, 1_047.790_2, 8.39),
                (1, 1_056.18, 8.39),
            ],
            24.78,
        );
        syllable.x = 1_039.4;
        assert!(
            (syllable.glyphs[1].x - (syllable.glyphs[0].x + syllable.glyphs[0].advance)).abs()
                > 0.0001
        );
        assert!(
            fitted_pieces(&syllable, 0.0, per_unit(8.0)).is_none(),
            "{:?}",
            syllable.glyphs
        );
    }

    #[test]
    fn a_run_whose_clusters_run_backwards_paints_as_laid_out() {
        let mut backwards = run(
            "\u{5d0}\u{5d1}\u{5d2}",
            &[(2, 40.0, 8.0), (1, 49.0, 8.0), (0, 58.0, 8.0)],
            27.0,
        );
        backwards.x = 40.0;
        assert!(fitted_pieces(&backwards, 0.0, per_unit(8.0)).is_none());
    }

    #[test]
    fn a_visually_right_to_left_run_paints_as_laid_out() {
        assert!(
            fitted_pieces(
                &run(
                    "\u{5d0}\u{5d1}\u{5d2}",
                    &[(0, 60.0, 10.0), (1, 50.0, 10.0), (2, 40.0, 10.0)],
                    30.0
                ),
                0.0,
                per_unit(8.0)
            )
            .is_none()
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
