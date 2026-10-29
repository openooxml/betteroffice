//! Shared tiny-skia raster helpers.

mod glyph;

pub use glyph::{CachedGlyph, GlyphCache, fill_glyph};

use ooxml_drawingml::GeometryPathCommand;
use ooxml_text::PathCmd;
use tiny_skia::{Color, ColorU8, Path, PathBuilder, Pixmap};

/// Encodes a premultiplied pixmap as RGBA PNG.
pub fn encode_png(pixmap: Pixmap, width: u32, height: u32) -> Result<Vec<u8>, String> {
    let pixels = pixmap.take_demultiplied();
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().map_err(|error| error.to_string())?;
        writer
            .write_image_data(&pixels)
            .map_err(|error| error.to_string())?;
    }
    Ok(data)
}

/// Parses hexadecimal RGB or RGBA digits, including short forms.
pub fn parse_hex_rgba(hex: &str) -> Option<[u8; 4]> {
    let expanded;
    let hex = match hex.len() {
        3 | 4 => {
            expanded = hex
                .chars()
                .flat_map(|character| [character, character])
                .collect::<String>();
            expanded.as_str()
        }
        6 | 8 => hex,
        _ => return None,
    };
    let byte = |index: usize| u8::from_str_radix(hex.get(index..index + 2)?, 16).ok();
    Some([
        byte(0)?,
        byte(2)?,
        byte(4)?,
        if hex.len() == 8 { byte(6)? } else { 255 },
    ])
}

/// Parses a hexadecimal color without its leading `#`.
pub fn parse_hex_color(hex: &str) -> Option<Color> {
    let [red, green, blue, alpha] = parse_hex_rgba(hex)?;
    Some(Color::from_rgba8(red, green, blue, alpha))
}

/// Parses hex, rgb(), rgba(), or transparent colors.
pub fn parse_css_color(value: &str) -> Result<Color, String> {
    if value.eq_ignore_ascii_case("transparent") {
        return Ok(Color::TRANSPARENT);
    }
    if let Some(hex) = value.strip_prefix('#') {
        return parse_hex_color(hex).ok_or_else(|| format!("bad color: {value}"));
    }
    if let Some(body) = value
        .strip_prefix("rgb(")
        .and_then(|value| value.strip_suffix(')'))
    {
        return parse_rgb_function(body, false).ok_or_else(|| format!("bad color: {value}"));
    }
    if let Some(body) = value
        .strip_prefix("rgba(")
        .and_then(|value| value.strip_suffix(')'))
    {
        return parse_rgb_function(body, true).ok_or_else(|| format!("bad color: {value}"));
    }
    Err(format!("bad color: {value}"))
}

fn parse_rgb_function(body: &str, alpha: bool) -> Option<Color> {
    let values = body.split(',').map(str::trim).collect::<Vec<_>>();
    if values.len() != if alpha { 4 } else { 3 } {
        return None;
    }
    let channel = |value: &str| -> Option<u8> {
        if let Some(percent) = value.strip_suffix('%') {
            let value = percent.parse::<f32>().ok()?.clamp(0.0, 100.0);
            Some((value * 2.55).round() as u8)
        } else {
            Some(value.parse::<f32>().ok()?.clamp(0.0, 255.0).round() as u8)
        }
    };
    let alpha = if alpha {
        if let Some(percent) = values[3].strip_suffix('%') {
            (percent.parse::<f32>().ok()?.clamp(0.0, 100.0) * 2.55).round() as u8
        } else {
            (values[3].parse::<f32>().ok()?.clamp(0.0, 1.0) * 255.0).round() as u8
        }
    } else {
        255
    };
    Some(Color::from_rgba8(
        channel(values[0])?,
        channel(values[1])?,
        channel(values[2])?,
        alpha,
    ))
}

/// Builds a geometry path with caller-defined point mapping.
pub fn path_from_commands(
    commands: &[GeometryPathCommand],
    mut map: impl FnMut(f64, f64) -> (f32, f32),
) -> Option<Path> {
    match try_path_from_commands(commands, |x, y| {
        Ok::<_, std::convert::Infallible>(map(x, y))
    }) {
        Ok(path) => path,
        Err(error) => match error {},
    }
}

/// Builds a geometry path with fallible point mapping.
pub fn try_path_from_commands<E>(
    commands: &[GeometryPathCommand],
    mut map: impl FnMut(f64, f64) -> Result<(f32, f32), E>,
) -> Result<Option<Path>, E> {
    let mut builder = PathBuilder::new();
    for command in commands {
        match *command {
            GeometryPathCommand::Move { x, y } => {
                let (x, y) = map(x, y)?;
                builder.move_to(x, y);
            }
            GeometryPathCommand::Line { x, y } => {
                let (x, y) = map(x, y)?;
                builder.line_to(x, y);
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                let (cpx, cpy) = map(cpx, cpy)?;
                let (x, y) = map(x, y)?;
                builder.quad_to(cpx, cpy, x, y);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let (cp1x, cp1y) = map(cp1x, cp1y)?;
                let (cp2x, cp2y) = map(cp2x, cp2y)?;
                let (x, y) = map(x, y)?;
                builder.cubic_to(cp1x, cp1y, cp2x, cp2y, x, y);
            }
            GeometryPathCommand::Close => builder.close(),
        }
    }
    Ok(builder.finish())
}

/// Builds a glyph outline without applying a transform.
pub fn outline_path(cmds: &[PathCmd]) -> Option<Path> {
    let mut builder = PathBuilder::new();
    for cmd in cmds {
        match *cmd {
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
    builder.finish()
}

/// Premultiplies complete RGBA pixels in place.
pub fn premultiply_rgba(data: &mut [u8]) {
    for pixel in data.as_chunks_mut::<4>().0 {
        let color = ColorU8::from_rgba(pixel[0], pixel[1], pixel[2], pixel[3]).premultiply();
        *pixel = [color.red(), color.green(), color.blue(), color.alpha()];
    }
}

/// A raster clip rectangle.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct RasterRect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

/// Intersects a clip with the target surface.
pub fn clipped_to_target(clip: RasterRect, width: u32, height: u32) -> Option<RasterRect> {
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
        .then_some(RasterRect {
            x: left,
            y: top,
            w: right - left,
            h: bottom - top,
        })
}

#[cfg(test)]
mod tests;
