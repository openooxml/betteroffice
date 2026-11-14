//! Shared coordinate and color helpers for the PDF writer.

use serde_json::Number;

use crate::PdfError;

/// CSS pixels per PDF point (96 dpi / 72 dpi).
pub(crate) const PX_TO_PT: f32 = 72.0 / 96.0;

pub(crate) fn points(value: &Number, what: &'static str) -> Result<f32, PdfError> {
    let value = value.as_f64().ok_or(PdfError::InvalidNumber(what))?;
    let value = (value * f64::from(PX_TO_PT)) as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

pub(crate) fn f64_to_points(value: f64, what: &'static str) -> Result<f32, PdfError> {
    let value = (value * f64::from(PX_TO_PT)) as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

/// Best-effort CSS font-size parser for a fallback `text` primitive.
pub(crate) fn font_size_px(css: &str) -> f32 {
    css.split_whitespace()
        .find_map(|part| part.strip_suffix("px"))
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(12.0)
}

/// Parse `#rgb`, `#rrggbb`, `rgb(...)` or `rgba(...)`; alpha is ignored.
pub(crate) fn parse_color(value: &str) -> Option<[f32; 3]> {
    let value = value.trim();
    if let Some(hex) = value.strip_prefix('#') {
        return match hex.len() {
            3 => Some([
                hex_channel(hex.get(0..1)?)?,
                hex_channel(hex.get(1..2)?)?,
                hex_channel(hex.get(2..3)?)?,
            ]),
            6 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
            ]),
            8 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
            ]),
            _ => None,
        };
    }
    let inner = value
        .strip_prefix("rgb(")
        .or_else(|| value.strip_prefix("rgba("))
        .and_then(|value| value.strip_suffix(')'))?;
    let mut parts = inner.split(',').map(str::trim);
    let r = channel(parts.next()?)?;
    let g = channel(parts.next()?)?;
    let b = channel(parts.next()?)?;
    Some([r, g, b])
}

fn hex_channel(value: &str) -> Option<f32> {
    let value = u8::from_str_radix(value, 16).ok()?;
    Some(f32::from(value) / 255.0)
}

fn channel(value: &str) -> Option<f32> {
    let value = value.parse::<f32>().ok()?;
    value.is_finite().then(|| (value / 255.0).clamp(0.0, 1.0))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_non_ascii_hex_colors() {
        assert_eq!(parse_color("#中"), None);
        assert_eq!(parse_color("#éa"), None);
        assert_eq!(parse_color("#12"), None);
        assert_eq!(parse_color("rgb(NaN, 0, 0)"), None);
        assert_eq!(parse_color("rgba(0, inf, 0, 1)"), None);
        assert!(parse_color("#112233").is_some());
    }
}
