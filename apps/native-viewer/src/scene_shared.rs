use std::collections::BTreeMap;

use vello::Scene;
use vello::kurbo::{Affine, Line, Rect, Stroke};
use vello::peniko::{Color, Fill};

#[derive(Default)]
pub struct SkipStats {
    pub counts: BTreeMap<String, usize>,
    pub reasons: BTreeMap<String, usize>,
}

impl SkipStats {
    pub fn total(&self) -> usize {
        self.counts.values().sum()
    }

    pub fn record(&mut self, kind: &str, reason: String) {
        *self.counts.entry(kind.to_owned()).or_default() += 1;
        *self.reasons.entry(reason).or_default() += 1;
    }
}

pub struct PageScene {
    pub background: Scene,
    pub scene: Scene,
    pub width: f64,
    pub height: f64,
    pub skipped: SkipStats,
}

pub fn with_clip_layers<T>(
    scene: &mut Scene,
    layers: &[(Affine, Rect)],
    draw: impl FnOnce(&mut Scene) -> Result<T, String>,
) -> Result<T, String> {
    for (transform, rect) in layers {
        scene.push_clip_layer(Fill::NonZero, *transform, rect);
    }
    let result = draw(scene);
    for _ in layers {
        scene.pop_layer();
    }
    result
}

pub fn with_dashes<I>(stroke: Stroke, dashes: I) -> Stroke
where
    I: IntoIterator<Item = f64>,
{
    stroke.with_dashes(0.0, dashes)
}

pub fn draw_placeholder(scene: &mut Scene, bounds: Option<Rect>) {
    draw_transformed_placeholder(scene, bounds, Affine::IDENTITY);
}

pub fn draw_transformed_placeholder(scene: &mut Scene, bounds: Option<Rect>, transform: Affine) {
    let mut rect = bounds.unwrap_or_else(|| Rect::new(0.0, 0.0, 18.0, 18.0));
    if rect.width() < 8.0 {
        rect = Rect::new(
            rect.x0 - 4.0,
            rect.y0,
            rect.x0 + 4.0,
            rect.y1.max(rect.y0 + 8.0),
        );
    }
    if rect.height() < 8.0 {
        rect = Rect::new(rect.x0, rect.y0 - 4.0, rect.x1, rect.y0 + 4.0);
    }
    scene.fill(
        Fill::NonZero,
        transform,
        Color::from_rgba8(255, 0, 255, 48),
        None,
        &rect,
    );
    let stroke = Stroke::new(1.5);
    let magenta = Color::from_rgba8(255, 0, 255, 255);
    scene.stroke(&stroke, transform, magenta, None, &rect);
    scene.stroke(
        &stroke,
        transform,
        magenta,
        None,
        &Line::new((rect.x0, rect.y0), (rect.x1, rect.y1)),
    );
    scene.stroke(
        &stroke,
        transform,
        magenta,
        None,
        &Line::new((rect.x1, rect.y0), (rect.x0, rect.y1)),
    );
}

pub fn color(value: &str, opacity: f32) -> Result<Color, String> {
    let (red, green, blue, alpha) = parse_color(value)?;
    let alpha = (f32::from(alpha) * opacity.clamp(0.0, 1.0)).round() as u8;
    Ok(Color::from_rgba8(red, green, blue, alpha))
}

pub fn strict_color(value: &str) -> Result<Color, String> {
    let hex = value
        .strip_prefix('#')
        .ok_or_else(|| format!("invalid XLSX color {value}"))?;
    if hex.len() != 6 && hex.len() != 8 {
        return Err(format!("invalid XLSX color {value}"));
    }
    let (red, green, blue, alpha) = parse_hex(value, hex)?;
    Ok(Color::from_rgba8(red, green, blue, alpha))
}

fn parse_color(value: &str) -> Result<(u8, u8, u8, u8), String> {
    if value.eq_ignore_ascii_case("transparent") {
        return Ok((0, 0, 0, 0));
    }
    if let Some(hex) = value.strip_prefix('#') {
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
            _ => return Err(format!("invalid color {value}")),
        };
        return parse_hex(value, hex);
    }
    for (prefix, alpha) in [("rgba(", true), ("rgb(", false)] {
        if let Some(body) = value
            .strip_prefix(prefix)
            .and_then(|body| body.strip_suffix(')'))
        {
            let parts = body.split(',').map(str::trim).collect::<Vec<_>>();
            if parts.len() != if alpha { 4 } else { 3 } {
                break;
            }
            let channel = |value: &str| -> Option<u8> {
                if let Some(percent) = value.strip_suffix('%') {
                    Some((percent.parse::<f32>().ok()?.clamp(0.0, 100.0) * 2.55).round() as u8)
                } else {
                    Some(value.parse::<f32>().ok()?.clamp(0.0, 255.0).round() as u8)
                }
            };
            let alpha = if alpha {
                if let Some(percent) = parts[3].strip_suffix('%') {
                    (percent
                        .parse::<f32>()
                        .map_err(|_| format!("invalid color {value}"))?
                        .clamp(0.0, 100.0)
                        * 2.55)
                        .round() as u8
                } else {
                    (parts[3]
                        .parse::<f32>()
                        .map_err(|_| format!("invalid color {value}"))?
                        .clamp(0.0, 1.0)
                        * 255.0)
                        .round() as u8
                }
            } else {
                255
            };
            return Ok((
                channel(parts[0]).ok_or_else(|| format!("invalid color {value}"))?,
                channel(parts[1]).ok_or_else(|| format!("invalid color {value}"))?,
                channel(parts[2]).ok_or_else(|| format!("invalid color {value}"))?,
                alpha,
            ));
        }
    }
    Err(format!("invalid color {value}"))
}

fn parse_hex(value: &str, hex: &str) -> Result<(u8, u8, u8, u8), String> {
    let byte = |index: usize| {
        hex.as_bytes()
            .get(index..index + 2)
            .and_then(|pair| std::str::from_utf8(pair).ok())
            .and_then(|pair| u8::from_str_radix(pair, 16).ok())
            .ok_or_else(|| format!("invalid color {value}"))
    };
    Ok((
        byte(0)?,
        byte(2)?,
        byte(4)?,
        if hex.len() == 8 { byte(6)? } else { 255 },
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_display_list_colors() {
        assert_eq!(parse_color("#369"), Ok((51, 102, 153, 255)));
        assert_eq!(parse_color("rgba(10, 20, 30, 0.5)"), Ok((10, 20, 30, 128)));
        assert!(parse_color("navy").is_err());
    }

    #[test]
    fn xlsx_colors_are_strict() {
        assert_eq!(
            strict_color("#33669980"),
            Ok(Color::from_rgba8(51, 102, 153, 128))
        );
        assert!(strict_color("#369").is_err());
        assert!(strict_color("red").is_err());
    }
}
