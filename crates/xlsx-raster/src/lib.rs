//! native raster backend: paints a display list to a png via tiny-skia.
//! server-side twin of the browser's canvas backend; never enters the wasm build.

mod font;

pub use font::measure_text;

use tiny_skia::{Color, Paint, PathBuilder, Pixmap, Rect, Stroke, StrokeDash, Transform};

use xlsx_render::{DisplayList, DrawCmd};

/// paint a display list and encode it as png bytes.
pub fn render_png(dl: &DisplayList) -> Result<Vec<u8>, String> {
    let w = (dl.width.ceil() as u32).max(1);
    let h = (dl.height.ceil() as u32).max(1);
    let mut pixmap = Pixmap::new(w, h).ok_or_else(|| "invalid pixmap size".to_string())?;
    pixmap.fill(Color::WHITE);

    for cmd in &dl.commands {
        match cmd {
            DrawCmd::FillRect { x, y, w, h, color } => {
                let Some(rect) = Rect::from_xywh(*x, *y, *w, *h) else {
                    continue;
                };
                let mut paint = Paint::default();
                paint.set_color(parse_color(color)?);
                paint.anti_alias = true;
                pixmap.fill_rect(rect, &paint, Transform::identity(), None);
            }
            DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color,
                style,
            } => {
                paint_line(&mut pixmap, *x1, *y1, *x2, *y2, *width, color, style)?;
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
                font_family: _,
            } => {
                font::paint_text(
                    &mut pixmap,
                    &font::TextRun {
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
                    },
                );
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

/// paint a line honoring its dash/double style; `double` approximates excel's
/// double border with two thin parallel passes.
#[allow(clippy::too_many_arguments)]
fn paint_line(
    pixmap: &mut Pixmap,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    color: &str,
    style: &Option<String>,
) -> Result<(), String> {
    let color = parse_color(color)?;
    match style.as_deref() {
        Some("double") => {
            let off = (width * 0.8).max(0.8);
            let horizontal = (y1 - y2).abs() <= (x1 - x2).abs();
            let (dx, dy) = if horizontal { (0.0, off) } else { (off, 0.0) };
            let w = (width * 0.6).max(0.5);
            stroke_seg(pixmap, x1 - dx, y1 - dy, x2 - dx, y2 - dy, w, color, None);
            stroke_seg(pixmap, x1 + dx, y1 + dy, x2 + dx, y2 + dy, w, color, None);
        }
        Some("dashed") => {
            let dash = StrokeDash::new(vec![4.0, 2.0], 0.0);
            stroke_seg(pixmap, x1, y1, x2, y2, width, color, dash);
        }
        Some("dotted") => {
            let dash = StrokeDash::new(vec![1.0, 2.0], 0.0);
            stroke_seg(pixmap, x1, y1, x2, y2, width, color, dash);
        }
        _ => stroke_seg(pixmap, x1, y1, x2, y2, width, color, None),
    }
    Ok(())
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
    pixmap.stroke_path(&path, &paint, &stroke, Transform::identity(), None);
}

/// parse a `#rrggbb` string into a tiny-skia color.
fn parse_color(s: &str) -> Result<Color, String> {
    let hex = s
        .strip_prefix('#')
        .ok_or_else(|| format!("bad color: {s}"))?;
    if hex.len() != 6 {
        return Err(format!("bad color: {s}"));
    }
    let byte =
        |i: usize| u8::from_str_radix(&hex[i..i + 2], 16).map_err(|_| format!("bad color: {s}"));
    Ok(Color::from_rgba8(byte(0)?, byte(2)?, byte(4)?, 255))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_render::Rect as DlRect;

    #[test]
    fn renders_png_with_magic_bytes() {
        let dl = DisplayList {
            width: 40.0,
            height: 20.0,
            commands: vec![
                DrawCmd::FillRect {
                    x: 0.0,
                    y: 0.0,
                    w: 40.0,
                    h: 20.0,
                    color: "#ffffff".into(),
                },
                DrawCmd::Line {
                    x1: 0.0,
                    y1: 10.0,
                    x2: 40.0,
                    y2: 10.0,
                    width: 1.0,
                    color: "#d4d4d4".into(),
                    style: None,
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
                    font_family: None,
                },
            ],
            grid: xlsx_render::GridMeta::default(),
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
            width: 10.0,
            height: 10.0,
            commands: vec![DrawCmd::FillRect {
                x: 0.0,
                y: 0.0,
                w: 10.0,
                h: 10.0,
                color: "red".into(),
            }],
            grid: xlsx_render::GridMeta::default(),
        };
        assert!(render_png(&dl).is_err());
    }
}
