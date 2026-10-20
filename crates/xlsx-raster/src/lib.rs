//! native raster backend: paints a display list to a png via tiny-skia.
//! server-side twin of the browser's canvas backend; never enters the wasm build.

mod font;
use std::collections::{HashMap, hash_map::Entry};

pub use font::measure_text;

use tiny_skia::{
    Color, FillRule, Mask, Paint, Path, PathBuilder, Pixmap, PixmapPaint, Rect, Stroke, StrokeDash,
    Transform,
};

use xlsx_render::{Align, DisplayList, DrawCmd, GeometryPathCommand, Rect as DlRect};

/// paint a display list and encode it as png bytes.
pub fn render_png(dl: &DisplayList) -> Result<Vec<u8>, String> {
    let w = (dl.width.ceil() as u32).max(1);
    let h = (dl.height.ceil() as u32).max(1);
    let mut pixmap = Pixmap::new(w, h).ok_or_else(|| "invalid pixmap size".to_string())?;
    pixmap.fill(Color::WHITE);
    let mut clip_surface = ClipSurface::default();
    let mut patterns = HashMap::new();

    for cmd in &dl.commands {
        match cmd {
            DrawCmd::PatternFill {
                x,
                y,
                w,
                h,
                pattern,
                fg,
                bg,
                origin_x,
                origin_y,
                clip,
            } => {
                let tile = match patterns.entry((*pattern, fg.clone(), bg.clone())) {
                    Entry::Occupied(entry) => entry.into_mut(),
                    Entry::Vacant(entry) => {
                        let mut tile =
                            Pixmap::new(8, 8).ok_or_else(|| "invalid pattern tile".to_string())?;
                        tile.fill(parse_color(bg)?);
                        let foreground = parse_color(fg)?.to_color_u8().premultiply();
                        for (y, row) in pattern.mask().into_iter().enumerate() {
                            for x in 0..8 {
                                if row & (1 << x) != 0 {
                                    tile.pixels_mut()[y * 8 + x] = foreground;
                                }
                            }
                        }
                        entry.insert(tile)
                    }
                };
                let Some(rect) = Rect::from_xywh(
                    pixel_round(*x),
                    pixel_round(*y),
                    pixel_round(x + w) - pixel_round(*x),
                    pixel_round(y + h) - pixel_round(*y),
                ) else {
                    continue;
                };
                let paint = Paint {
                    shader: tiny_skia::Pattern::new(
                        tile.as_ref(),
                        tiny_skia::SpreadMode::Repeat,
                        tiny_skia::FilterQuality::Nearest,
                        1.0,
                        Transform::from_translate(pixel_round(*origin_x), pixel_round(*origin_y)),
                    ),
                    anti_alias: false,
                    ..Paint::default()
                };
                if let Some(clip) = clip {
                    clip_surface.paint(&mut pixmap, *clip, |target, transform, mask| {
                        target.fill_rect(rect, &paint, transform, mask);
                        Ok(())
                    })?;
                } else {
                    pixmap.fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
            DrawCmd::FillRect {
                x,
                y,
                w,
                h,
                color,
                clip,
            } => {
                let rect = Rect::from_xywh(*x, *y, *w, *h)
                    .ok_or_else(|| "invalid fill rectangle".to_string())?;
                let mut paint = Paint::default();
                paint.set_color(parse_color(color)?);
                paint.anti_alias = true;
                if let Some(clip) = clip {
                    clip_surface.paint(&mut pixmap, *clip, |target, transform, mask| {
                        target.fill_rect(rect, &paint, transform, mask);
                        Ok(())
                    })?;
                } else {
                    pixmap.fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
            DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color,
                style,
                rail,
                clip,
            } => {
                let paint =
                    |target: &mut Pixmap, transform: Transform, mask: Option<&Mask>| match rail {
                        Some(rail) if style.as_deref() == Some("double") => paint_double(
                            target, *x1, *y1, *x2, *y2, *width, *rail, color, transform, mask,
                        ),
                        _ => paint_line(
                            target, *x1, *y1, *x2, *y2, *width, color, style, transform, mask,
                        ),
                    };
                if let Some(clip) = clip {
                    clip_surface.paint(&mut pixmap, *clip, paint)?;
                } else {
                    paint(&mut pixmap, Transform::identity(), None)?;
                }
            }
            DrawCmd::Path {
                commands,
                fill,
                stroke,
                clip,
            } => {
                let path = build_path(commands)?;
                let mut paint = Paint::default();
                paint.set_color(parse_color(fill)?);
                paint.anti_alias = true;
                let stroke = if let Some(stroke) = stroke {
                    let mut paint = Paint::default();
                    paint.set_color(parse_color(&stroke.color)?);
                    paint.anti_alias = true;
                    Some((paint, stroke.width))
                } else {
                    None
                };
                let paint_path =
                    |target: &mut Pixmap, transform: Transform, mask: Option<&Mask>| {
                        target.fill_path(&path, &paint, FillRule::Winding, transform, mask);
                        if let Some((stroke_paint, width)) = &stroke {
                            target.stroke_path(
                                &path,
                                stroke_paint,
                                &Stroke {
                                    width: *width,
                                    ..Stroke::default()
                                },
                                transform,
                                mask,
                            );
                        }
                        Ok(())
                    };
                if let Some(clip) = clip {
                    clip_surface.paint(&mut pixmap, *clip, paint_path)?;
                } else {
                    pixmap.fill_path(
                        &path,
                        &paint,
                        FillRule::Winding,
                        Transform::identity(),
                        None,
                    );
                    if let Some((stroke_paint, width)) = &stroke {
                        pixmap.stroke_path(
                            &path,
                            stroke_paint,
                            &Stroke {
                                width: *width,
                                ..Stroke::default()
                            },
                            Transform::identity(),
                            None,
                        );
                    }
                }
            }
            DrawCmd::Text {
                x,
                y,
                text,
                font_size,
                color,
                clip,
                align,
                bold,
                italic,
                underline,
                strike,
                highlight,
                dashed_underline,
                font_family: _,
                number_format,
                ghost: _,
                chart: _,
            } => {
                let run = font::TextRun {
                    x: *x,
                    y: *y,
                    text,
                    font_size_pt: *font_size,
                    color: parse_color(color)?,
                    align: *align,
                    clip,
                    bold: *bold,
                    italic: *italic,
                    underline: *underline,
                    strike: *strike,
                    highlight: highlight.as_deref().map(parse_color).transpose()?,
                    dashed_underline: *dashed_underline,
                };
                if let Some(layout) = number_format {
                    xlsx_render::paint_number_format(
                        text,
                        layout,
                        *align,
                        |literal| font::measure_text(literal, *font_size),
                        |literal, x| {
                            font::paint_text(
                                &mut pixmap,
                                &font::TextRun {
                                    x,
                                    text: literal,
                                    align: Align::Left,
                                    ..run
                                },
                            )
                        },
                    );
                } else {
                    font::paint_text(&mut pixmap, &run);
                }
            }
        }
    }

    let pixels = pixmap.take_demultiplied();
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, w, h);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().map_err(|e| e.to_string())?;
        writer
            .write_image_data(&pixels)
            .map_err(|e| e.to_string())?;
    }
    Ok(data)
}

/// paint a line honoring its dash/double style; a `double` without a rail width
/// keeps the legacy approximation: two thin passes around the unsnapped line.
#[allow(clippy::too_many_arguments)]
fn paint_line(
    pixmap: &mut Pixmap,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    color: &str,
    style: &Option<std::sync::Arc<str>>,
    transform: Transform,
    mask: Option<&Mask>,
) -> Result<(), String> {
    let color = parse_color(color)?;
    match style.as_deref() {
        Some("double") => {
            let off = (width * 0.8).max(0.8);
            let horizontal = (y1 - y2).abs() <= (x1 - x2).abs();
            let (dx, dy) = if horizontal { (0.0, off) } else { (off, 0.0) };
            let w = (width * 0.6).max(0.5);
            stroke_seg(
                pixmap,
                x1 - dx,
                y1 - dy,
                x2 - dx,
                y2 - dy,
                w,
                color,
                None,
                transform,
                mask,
            );
            stroke_seg(
                pixmap,
                x1 + dx,
                y1 + dy,
                x2 + dx,
                y2 + dy,
                w,
                color,
                None,
                transform,
                mask,
            );
        }
        Some(style) if !xlsx_render::border_dash(style).is_empty() => {
            let pen = width.round().max(1.0);
            let (x1, y1, x2, y2) = snap(x1, y1, x2, y2, pen);
            let scale = if style.starts_with("medium") {
                width / 2.0
            } else {
                width
            };
            let dash = StrokeDash::new(
                xlsx_render::border_dash(style)
                    .iter()
                    .map(|run| run * scale)
                    .collect(),
                0.0,
            );
            stroke_seg(pixmap, x1, y1, x2, y2, pen, color, dash, transform, mask);
        }
        _ => stroke_seg(pixmap, x1, y1, x2, y2, width, color, None, transform, mask),
    }
    Ok(())
}

/// paint a double line as two `rail`-wide passes on the outer edges of `width`.
#[allow(clippy::too_many_arguments)]
fn paint_double(
    pixmap: &mut Pixmap,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    rail: f32,
    color: &str,
    transform: Transform,
    mask: Option<&Mask>,
) -> Result<(), String> {
    let color = parse_color(color)?;
    let (x1, y1, x2, y2) = snap(x1, y1, x2, y2, width.round().max(1.0));
    let off = (width - rail) / 2.0;
    let (dx, dy) = if (y1 - y2).abs() <= (x1 - x2).abs() {
        (0.0, off)
    } else {
        (off, 0.0)
    };
    for sign in [-1.0, 1.0] {
        stroke_seg(
            pixmap,
            x1 + sign * dx,
            y1 + sign * dy,
            x2 + sign * dx,
            y2 + sign * dy,
            rail,
            color,
            None,
            transform,
            mask,
        );
    }
    Ok(())
}

/// centre an axis-aligned line on device pixels for a pen `pen` pixels wide.
fn snap(x1: f32, y1: f32, x2: f32, y2: f32, pen: f32) -> (f32, f32, f32, f32) {
    let offset = (pen as u32 % 2) as f32 / 2.0;
    if y1 == y2 {
        let y = y1.floor() + offset;
        (x1.round(), y, x2.round(), y)
    } else if x1 == x2 {
        let x = x1.floor() + offset;
        (x, y1.round(), x, y2.round())
    } else {
        (x1, y1, x2, y2)
    }
}

/// stroke a single segment with an optional dash pattern.
#[allow(clippy::too_many_arguments)]
fn stroke_seg(
    pixmap: &mut Pixmap,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    color: Color,
    dash: Option<StrokeDash>,
    transform: Transform,
    mask: Option<&Mask>,
) {
    let mut pb = PathBuilder::new();
    pb.move_to(x1, y1);
    pb.line_to(x2, y2);
    let Some(path) = pb.finish() else {
        return;
    };
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    let stroke = Stroke {
        width,
        dash,
        ..Stroke::default()
    };
    pixmap.stroke_path(&path, &paint, &stroke, transform, mask);
}

#[derive(Default)]
struct ClipSurface {
    rect: Option<DlRect>,
    origin: (i32, i32),
    pixmap: Option<Pixmap>,
    mask: Option<Mask>,
}

fn pixel_round(value: f32) -> f32 {
    (value + 0.5).floor()
}

impl ClipSurface {
    fn paint<F>(&mut self, target: &mut Pixmap, clip: DlRect, painter: F) -> Result<(), String>
    where
        F: FnOnce(&mut Pixmap, Transform, Option<&Mask>) -> Result<(), String>,
    {
        let Some(clip) = clipped_to_target(clip, target.width(), target.height()) else {
            return Ok(());
        };
        if self.rect != Some(clip) {
            let origin_x = clip.x.floor() as i32;
            let origin_y = clip.y.floor() as i32;
            let width = ((clip.x + clip.w).ceil() as i32 - origin_x) as u32;
            let height = ((clip.y + clip.h).ceil() as i32 - origin_y) as u32;
            let mut mask =
                Mask::new(width, height).ok_or_else(|| "invalid clip mask size".to_string())?;
            let rect = Rect::from_xywh(
                clip.x - origin_x as f32,
                clip.y - origin_y as f32,
                clip.w,
                clip.h,
            )
            .ok_or_else(|| "invalid clip rectangle".to_string())?;
            let path = PathBuilder::from_rect(rect);
            mask.fill_path(&path, FillRule::Winding, true, Transform::identity());
            self.rect = Some(clip);
            self.origin = (origin_x, origin_y);
            self.pixmap =
                Some(Pixmap::new(width, height).ok_or_else(|| "invalid clip size".to_string())?);
            self.mask = Some(mask);
        }
        let pixmap = self
            .pixmap
            .as_mut()
            .ok_or_else(|| "clip surface is unavailable".to_string())?;
        pixmap.fill(Color::TRANSPARENT);
        let transform = Transform::from_translate(-(self.origin.0 as f32), -(self.origin.1 as f32));
        painter(pixmap, transform, self.mask.as_ref())?;
        target.draw_pixmap(
            self.origin.0,
            self.origin.1,
            pixmap.as_ref(),
            &PixmapPaint::default(),
            Transform::identity(),
            None,
        );
        Ok(())
    }
}

fn clipped_to_target(clip: DlRect, width: u32, height: u32) -> Option<DlRect> {
    let left = clip.x.max(0.0);
    let top = clip.y.max(0.0);
    let right = (clip.x + clip.w).min(width as f32);
    let bottom = (clip.y + clip.h).min(height as f32);
    (left.is_finite()
        && top.is_finite()
        && right.is_finite()
        && bottom.is_finite()
        && right > left
        && bottom > top)
        .then_some(DlRect {
            x: left,
            y: top,
            w: right - left,
            h: bottom - top,
        })
}

fn build_path(commands: &[GeometryPathCommand]) -> Result<Path, String> {
    let mut builder = PathBuilder::new();
    for command in commands {
        match command {
            GeometryPathCommand::Move { x, y } => {
                builder.move_to(path_coord(*x)?, path_coord(*y)?);
            }
            GeometryPathCommand::Line { x, y } => {
                builder.line_to(path_coord(*x)?, path_coord(*y)?);
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                builder.quad_to(
                    path_coord(*cpx)?,
                    path_coord(*cpy)?,
                    path_coord(*x)?,
                    path_coord(*y)?,
                );
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                builder.cubic_to(
                    path_coord(*cp1x)?,
                    path_coord(*cp1y)?,
                    path_coord(*cp2x)?,
                    path_coord(*cp2y)?,
                    path_coord(*x)?,
                    path_coord(*y)?,
                );
            }
            GeometryPathCommand::Close => builder.close(),
        }
    }
    builder
        .finish()
        .ok_or_else(|| "path has no drawable commands".to_string())
}

fn path_coord(value: f64) -> Result<f32, String> {
    let value = value as f32;
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| "path contains a non-finite coordinate".to_string())
}

/// parse a `#rrggbb` string into a tiny-skia color.
fn parse_color(s: &str) -> Result<Color, String> {
    let hex = s
        .strip_prefix('#')
        .ok_or_else(|| format!("bad color: {s}"))?;
    if hex.len() != 6 && hex.len() != 8 {
        return Err(format!("bad color: {s}"));
    }
    let byte = |i: usize| {
        hex.as_bytes()
            .get(i..i + 2)
            .and_then(|pair| std::str::from_utf8(pair).ok())
            .and_then(|pair| u8::from_str_radix(pair, 16).ok())
            .ok_or_else(|| format!("bad color: {s}"))
    };
    Ok(Color::from_rgba8(
        byte(0)?,
        byte(2)?,
        byte(4)?,
        if hex.len() == 8 { byte(6)? } else { 255 },
    ))
}

#[cfg(test)]
mod tests {
    #[test]
    fn pattern_border_pixels_follow_device_mask_and_dash_runs() {
        let mut workbook = xlsx_model::Workbook::default();
        let mut sheet = xlsx_model::Sheet::new("Sheet1");
        sheet.col_widths.insert(0, 8.0 / 7.0);
        sheet.row_heights.insert(0, 6.0);
        let mut format = workbook.styles.cell_format(None);
        format.fill = serde_json::from_value(serde_json::json!({"Pattern": {
            "pattern": "lightUp", "fg": {"Rgb": "#000000"}, "bg": {"Rgb": "#ffffff"}
        }}))
        .unwrap();
        let style = workbook.styles.intern_cell_format(&format).unwrap();
        for col in 0..3 {
            sheet.set_cell(
                xlsx_model::CellRef::new(0, col),
                xlsx_model::Cell {
                    style,
                    ..Default::default()
                },
            );
        }
        workbook.sheets.push(sheet);
        let mut dl = xlsx_render::build_display_list(
            &workbook,
            xlsx_model::SheetId(0),
            &xlsx_render::Viewport {
                x: 0.0,
                y: 0.0,
                width: 32.0,
                height: 16.0,
            },
        )
        .unwrap();
        dl.commands
            .retain(|cmd| !matches!(cmd, DrawCmd::Line { .. }));
        dl.commands.push(DrawCmd::Line {
            x1: 0.0,
            y1: 12.5,
            x2: 32.0,
            y2: 12.5,
            width: 1.0,
            color: "#000000".into(),
            style: Some("dashDot".into()),
            rail: None,
            clip: None,
        });
        let pixels = Pixmap::decode_png(&render_png(&dl).unwrap()).unwrap();
        for y in 0..8 {
            for x in 0..32 {
                let expected = if (x + y) % 4 == 0 { 0 } else { 255 };
                assert_eq!(pixels.pixel(x, y).unwrap().red(), expected, "{x},{y}");
            }
        }
        for x in 0..32 {
            let expected = if matches!(x % 9, 0..=3 | 6) { 0 } else { 255 };
            assert_eq!(pixels.pixel(x, 12).unwrap().red(), expected, "dash at {x}");
        }
        let zoomed = xlsx_render::scaled(dl, 2.0);
        let pixels = Pixmap::decode_png(&render_png(&zoomed).unwrap()).unwrap();
        for y in 0..16 {
            for x in 0..64 {
                assert_eq!(
                    pixels.pixel(x, y).unwrap().red(),
                    if (x + y) % 4 == 0 { 0 } else { 255 },
                    "zoomed {x},{y}"
                );
            }
        }
    }

    #[test]
    fn pattern_border_all_dash_runs_and_double_pixels() {
        for (style, width, runs) in [
            ("dashed", 1.0, vec![4, 2]),
            ("dotted", 1.0, vec![1, 2]),
            ("mediumDashed", 2.0, vec![8, 4]),
            ("dashDot", 1.0, vec![4, 2, 1, 2]),
            ("mediumDashDot", 2.0, vec![8, 4, 2, 4]),
            ("dashDotDot", 1.0, vec![4, 2, 1, 2, 1, 2]),
            ("mediumDashDotDot", 2.0, vec![8, 4, 2, 4, 2, 4]),
            ("slantDashDot", 1.0, vec![4, 2, 1, 2]),
        ] {
            let mut pixmap = Pixmap::new(64, 8).unwrap();
            pixmap.fill(Color::WHITE);
            let y = if width == 1.0 { 4.5 } else { 4.0 };
            paint_line(
                &mut pixmap,
                0.0,
                y,
                64.0,
                y,
                width,
                "#000000",
                &Some(style.into()),
                Transform::identity(),
                None,
            )
            .unwrap();
            let period: u32 = runs.iter().sum();
            for x in 0..64 {
                let mut offset = x % period;
                let mut on = true;
                for run in &runs {
                    if offset < *run {
                        break;
                    }
                    offset -= *run;
                    on = !on;
                }
                assert_eq!(
                    pixmap.pixel(x, 4).unwrap().red(),
                    if on { 0 } else { 255 },
                    "{style} at {x}"
                );
            }
        }
        let mut pixmap = Pixmap::new(16, 8).unwrap();
        pixmap.fill(Color::WHITE);
        paint_double(
            &mut pixmap,
            0.0,
            4.5,
            16.0,
            4.5,
            3.0,
            1.0,
            "#000000",
            Transform::identity(),
            None,
        )
        .unwrap();
        assert_eq!(pixmap.pixel(8, 3).unwrap().red(), 0);
        assert_eq!(pixmap.pixel(8, 4).unwrap().red(), 255);
        assert_eq!(pixmap.pixel(8, 5).unwrap().red(), 0);
    }

    #[test]
    fn double_line_without_rail_keeps_released_rails() {
        let black = parse_color("#000000").unwrap();
        for width in [0.5_f32, 1.0, 2.0] {
            for (x1, y1, x2, y2) in [(0.0, 8.0, 24.0, 8.0), (12.0, 0.0, 12.0, 16.0)] {
                let blank = || {
                    let mut pixmap = Pixmap::new(24, 16).unwrap();
                    pixmap.fill(Color::WHITE);
                    pixmap
                };
                let mut actual = blank();
                paint_line(
                    &mut actual,
                    x1,
                    y1,
                    x2,
                    y2,
                    width,
                    "#000000",
                    &Some("double".into()),
                    Transform::identity(),
                    None,
                )
                .unwrap();
                let mut expected = blank();
                let (rail, offset) = ((width * 0.6).max(0.5), (width * 0.8).max(0.8));
                let (dx, dy) = if y1 == y2 {
                    (0.0, offset)
                } else {
                    (offset, 0.0)
                };
                for sign in [-1.0, 1.0] {
                    stroke_seg(
                        &mut expected,
                        x1 + sign * dx,
                        y1 + sign * dy,
                        x2 + sign * dx,
                        y2 + sign * dy,
                        rail,
                        black,
                        None,
                        Transform::identity(),
                        None,
                    );
                }
                assert!(actual.data() == expected.data(), "{width} {x1},{y1}");
            }
        }
    }

    #[test]
    fn produced_double_border_paints_rail_gap_rail() {
        let mut workbook = xlsx_model::Workbook::default();
        let mut sheet = xlsx_model::Sheet::new("Sheet1");
        let mut format = workbook.styles.cell_format(None);
        format.border.bottom = Some(xlsx_model::BorderEdge {
            style: xlsx_model::BorderStyle::Double,
            color: Some(xlsx_model::Color::Rgb("#ff0000".into())),
        });
        let style = workbook.styles.intern_cell_format(&format).unwrap();
        sheet.set_cell(
            xlsx_model::CellRef::new(0, 0),
            xlsx_model::Cell {
                style,
                ..Default::default()
            },
        );
        workbook.sheets.push(sheet);
        let dl = xlsx_render::build_display_list(
            &workbook,
            xlsx_model::SheetId(0),
            &xlsx_render::Viewport {
                x: 0.0,
                y: 0.0,
                width: 48.0,
                height: 40.0,
            },
        )
        .unwrap();
        for factor in [1.0, 2.0] {
            let pixels =
                Pixmap::decode_png(&render_png(&xlsx_render::scaled(dl.clone(), factor)).unwrap())
                    .unwrap();
            let ink: Vec<u32> = (0..pixels.height())
                .filter(|&y| {
                    let pixel = pixels.pixel(16, y).unwrap();
                    pixel.red() == 255 && pixel.green() == 0
                })
                .collect();
            let rail = factor as u32;
            assert_eq!(ink.len() as u32, 2 * rail, "{factor}: {ink:?}");
            assert_eq!(ink[rail as usize] - ink[0], 2 * rail, "{factor}: {ink:?}");
            for y in ink[0]..ink[0] + 3 * rail {
                let pixel = pixels.pixel(16, y).unwrap();
                assert!(
                    ink.contains(&y) || pixel.green() > 128,
                    "{factor}: gap at {y}"
                );
            }
        }
    }

    #[test]
    fn dashed_line_below_one_pixel_is_centred_on_its_clamped_width() {
        let mut pixmap = Pixmap::new(32, 8).unwrap();
        pixmap.fill(Color::WHITE);
        paint_line(
            &mut pixmap,
            0.0,
            4.0,
            32.0,
            4.0,
            0.25,
            "#000000",
            &Some("dashed".into()),
            Transform::identity(),
            None,
        )
        .unwrap();
        for x in 0..32 {
            assert_eq!(pixmap.pixel(x, 3).unwrap().red(), 255, "above at {x}");
            assert_eq!(pixmap.pixel(x, 5).unwrap().red(), 255, "below at {x}");
        }
        assert_eq!(pixmap.pixel(0, 4).unwrap().red(), 0);
    }
    use super::*;
    use serde::Deserialize;
    use tiny_skia::PathSegment;
    use xlsx_render::Rect as DlRect;

    #[derive(Deserialize)]
    struct AgreementFixture {
        commands: Vec<GeometryPathCommand>,
        trace: Vec<PathTrace>,
    }

    #[derive(Debug, Deserialize, PartialEq)]
    struct PathTrace {
        verb: String,
        points: Vec<f32>,
    }

    #[test]
    fn renders_png_with_magic_bytes() {
        let dl = DisplayList {
            truncated: false,
            width: 40.0,
            height: 20.0,
            commands: vec![
                DrawCmd::FillRect {
                    x: 0.0,
                    y: 0.0,
                    w: 40.0,
                    h: 20.0,
                    color: "#ffffff".into(),
                    clip: None,
                },
                DrawCmd::Line {
                    x1: 0.0,
                    y1: 10.0,
                    x2: 40.0,
                    y2: 10.0,
                    width: 1.0,
                    color: "#d4d4d4".into(),
                    style: None,
                    rail: None,
                    clip: None,
                },
                DrawCmd::Text {
                    x: 2.0,
                    y: 12.0,
                    text: "painted".into(),
                    font_size: 11.0,
                    color: "#000000".into(),
                    clip: DlRect {
                        x: 0.0,
                        y: 0.0,
                        w: 40.0,
                        h: 20.0,
                    },
                    align: xlsx_render::Align::Left,
                    bold: false,
                    italic: false,
                    underline: false,
                    strike: false,
                    highlight: None,
                    dashed_underline: false,
                    font_family: None,
                    number_format: None,
                    ghost: false,
                    chart: false,
                },
            ],
            grid: xlsx_render::GridMeta::default(),
            hyperlinks: Vec::new(),
            charts: Vec::new(),
        };

        let png = render_png(&dl).unwrap();
        assert!(png.len() > 8);
        assert_eq!(
            &png[0..8],
            &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]
        );
    }

    #[test]
    fn rejects_malformed_color() {
        let dl = DisplayList {
            truncated: false,
            width: 10.0,
            height: 10.0,
            commands: vec![DrawCmd::FillRect {
                x: 0.0,
                y: 0.0,
                w: 10.0,
                h: 10.0,
                color: "red".into(),
                clip: None,
            }],
            grid: xlsx_render::GridMeta::default(),
            hyperlinks: Vec::new(),
            charts: Vec::new(),
        };
        assert!(render_png(&dl).is_err());
    }

    #[test]
    fn path_backends_follow_shared_trace_contract() {
        let fixture: AgreementFixture = serde_json::from_str(include_str!(
            "../../../packages/xlsx/test-fixtures/path-backend-agreement.json"
        ))
        .unwrap();
        let path = build_path(&fixture.commands).unwrap();
        let trace = path
            .segments()
            .map(|segment| match segment {
                PathSegment::MoveTo(point) => PathTrace {
                    verb: "move".into(),
                    points: vec![point.x, point.y],
                },
                PathSegment::LineTo(point) => PathTrace {
                    verb: "line".into(),
                    points: vec![point.x, point.y],
                },
                PathSegment::QuadTo(control, point) => PathTrace {
                    verb: "quad".into(),
                    points: vec![control.x, control.y, point.x, point.y],
                },
                PathSegment::CubicTo(first, second, point) => PathTrace {
                    verb: "cubic".into(),
                    points: vec![first.x, first.y, second.x, second.y, point.x, point.y],
                },
                PathSegment::Close => PathTrace {
                    verb: "close".into(),
                    points: Vec::new(),
                },
            })
            .collect::<Vec<_>>();
        assert_eq!(trace, fixture.trace);
    }
}
