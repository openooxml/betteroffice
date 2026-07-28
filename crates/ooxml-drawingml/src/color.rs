use serde::{Deserialize, Serialize};

use crate::Theme;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColorValue {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto: Option<bool>,
    /// `a:lumMod`, as a fraction of the colour's own luminance.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub luminance_modulation: Option<f64>,
    /// `a:lumOff`, added to luminance after `a:lumMod`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub luminance_offset: Option<f64>,
    /// `a:satMod`, as a fraction of the colour's own saturation.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub saturation_modulation: Option<f64>,
    /// `a:alpha`, as a fraction. Opaque hex output drops it; use
    /// [`resolve_color_value_to_rgba_hex`] to keep it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alpha: Option<f64>,
}

pub fn parse_color_value(
    rgb: Option<&str>,
    theme_color: Option<&str>,
    theme_tint: Option<&str>,
    theme_shade: Option<&str>,
) -> ColorValue {
    ColorValue {
        rgb: rgb
            .filter(|value| !value.is_empty() && *value != "auto")
            .map(str::to_owned),
        auto: (rgb == Some("auto")).then_some(true),
        theme_color: theme_color
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        theme_tint: theme_tint
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        theme_shade: theme_shade
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        ..ColorValue::default()
    }
}

pub fn resolve_color_value_to_hex(color: Option<&ColorValue>) -> Option<String> {
    resolve_color_value_to_hex_with_theme(color, None)
}

pub fn resolve_color_value_to_hex_with_theme(
    color: Option<&ColorValue>,
    theme: Option<&Theme>,
) -> Option<String> {
    let color = color?;
    let rgb = color.rgb.as_deref().or_else(|| {
        color.theme_color.as_deref().map(|slot| {
            theme
                .and_then(|theme| theme.color_scheme.get(slot))
                .unwrap_or_else(|| default_theme_color(slot))
        })
    })?;
    let mut channels = parse_rgb(rgb)?;
    if let Some(shade) = color.theme_shade.as_deref().and_then(parse_modifier) {
        channels = channels.map(|channel| (f64::from(channel) * shade).round() as u8);
    }
    if let Some(tint) = color.theme_tint.as_deref().and_then(parse_modifier) {
        channels = channels.map(|channel| {
            (f64::from(channel) + (255.0 - f64::from(channel)) * tint).round() as u8
        });
    }
    channels = apply_hsl_modifiers(channels, color);
    Some(format!(
        "#{:02X}{:02X}{:02X}",
        channels[0], channels[1], channels[2]
    ))
}

/// `#RRGGBBAA`, for hosts that can paint `a:alpha`. Every other resolver
/// returns opaque `#RRGGBB`, which many consumers require.
pub fn resolve_color_value_to_rgba_hex(
    color: Option<&ColorValue>,
    theme: Option<&Theme>,
) -> Option<String> {
    let hex = resolve_color_value_to_hex_with_theme(color, theme)?;
    let alpha = color
        .and_then(|color| color.alpha)
        .filter(|alpha| alpha.is_finite() && *alpha < 1.0)
        .unwrap_or(1.0)
        .clamp(0.0, 1.0);
    Some(format!("{hex}{:02X}", (alpha * 255.0).round() as u8))
}

/// `a:satMod`, `a:lumMod` and `a:lumOff`, which scale and shift the colour in
/// HSL space rather than per channel.
fn apply_hsl_modifiers(channels: [u8; 3], color: &ColorValue) -> [u8; 3] {
    let modulation = |value: Option<f64>| value.filter(|value| value.is_finite() && *value >= 0.0);
    let saturation = modulation(color.saturation_modulation);
    let luminance = modulation(color.luminance_modulation);
    let offset = color
        .luminance_offset
        .filter(|value| value.is_finite() && *value != 0.0);
    if saturation.is_none() && luminance.is_none() && offset.is_none() {
        return channels;
    }
    let (hue, mut sat, mut lum) = rgb_to_hsl(channels);
    if let Some(saturation) = saturation {
        sat = (sat * saturation).clamp(0.0, 1.0);
    }
    if let Some(luminance) = luminance {
        lum = (lum * luminance).clamp(0.0, 1.0);
    }
    if let Some(offset) = offset {
        lum = (lum + offset).clamp(0.0, 1.0);
    }
    hsl_to_rgb(hue, sat, lum)
}

fn rgb_to_hsl(channels: [u8; 3]) -> (f64, f64, f64) {
    let [red, green, blue] = channels.map(|channel| f64::from(channel) / 255.0);
    let max = red.max(green).max(blue);
    let min = red.min(green).min(blue);
    let lum = (max + min) / 2.0;
    let span = max - min;
    if span <= f64::EPSILON {
        return (0.0, 0.0, lum);
    }
    let sat = span / (1.0 - (2.0 * lum - 1.0).abs());
    let hue = if max == red {
        ((green - blue) / span).rem_euclid(6.0)
    } else if max == green {
        (blue - red) / span + 2.0
    } else {
        (red - green) / span + 4.0
    };
    (hue * 60.0, sat.clamp(0.0, 1.0), lum)
}

fn hsl_to_rgb(hue: f64, sat: f64, lum: f64) -> [u8; 3] {
    let chroma = (1.0 - (2.0 * lum - 1.0).abs()) * sat;
    let sector = hue.rem_euclid(360.0) / 60.0;
    let second = chroma * (1.0 - (sector.rem_euclid(2.0) - 1.0).abs());
    let (red, green, blue) = match sector as u8 {
        0 => (chroma, second, 0.0),
        1 => (second, chroma, 0.0),
        2 => (0.0, chroma, second),
        3 => (0.0, second, chroma),
        4 => (second, 0.0, chroma),
        _ => (chroma, 0.0, second),
    };
    let base = lum - chroma / 2.0;
    [red, green, blue].map(|channel| ((channel + base).clamp(0.0, 1.0) * 255.0).round() as u8)
}

fn parse_rgb(value: &str) -> Option<[u8; 3]> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if value.len() != 6 {
        return None;
    }
    let packed = u32::from_str_radix(value, 16).ok()?;
    Some([
        ((packed >> 16) & 0xff) as u8,
        ((packed >> 8) & 0xff) as u8,
        (packed & 0xff) as u8,
    ])
}

fn parse_modifier(value: &str) -> Option<f64> {
    let byte = u8::from_str_radix(value, 16).ok()?;
    Some(f64::from(byte) / 255.0)
}

pub fn default_theme_color(slot: &str) -> &str {
    match slot {
        "dk1" | "text1" => "000000",
        "lt1" | "background1" => "FFFFFF",
        "dk2" | "text2" => "44546A",
        "lt2" | "background2" => "E7E6E6",
        "accent1" => "4472C4",
        "accent2" => "ED7D31",
        "accent3" => "A5A5A5",
        "accent4" => "FFC000",
        "accent5" => "5B9BD5",
        "accent6" => "70AD47",
        "hlink" => "0563C1",
        "folHlink" => "954F72",
        _ => "000000",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn luminance_and_saturation_modifiers_resolve_in_hsl() {
        let modified = |luminance: Option<f64>, offset: Option<f64>, saturation: Option<f64>| {
            resolve_color_value_to_hex(Some(&ColorValue {
                rgb: Some("4472C4".to_owned()),
                luminance_modulation: luminance,
                luminance_offset: offset,
                saturation_modulation: saturation,
                ..ColorValue::default()
            }))
            .expect("resolves")
        };
        assert_eq!(modified(None, None, None), "#4472C4");
        assert_eq!(modified(Some(1.0), None, None), "#4472C4");
        // The pairing Office writes for a 40% lighter accent.
        assert_eq!(modified(Some(0.6), Some(0.4), None), "#8FAADC");
        assert_eq!(modified(Some(0.75), None, None), "#2F5597");
        assert_eq!(modified(None, None, Some(0.0)), "#848484");
        assert_eq!(
            resolve_color_value_to_hex(Some(&ColorValue {
                rgb: Some("808080".to_owned()),
                saturation_modulation: Some(4.0),
                ..ColorValue::default()
            }))
            .as_deref(),
            Some("#808080"),
            "a grey has no saturation to modulate"
        );
    }

    #[test]
    fn alpha_reaches_the_rgba_resolver_and_never_the_opaque_one() {
        let translucent = ColorValue {
            rgb: Some("112233".to_owned()),
            alpha: Some(0.5),
            ..ColorValue::default()
        };
        assert_eq!(
            resolve_color_value_to_hex(Some(&translucent)).as_deref(),
            Some("#112233")
        );
        assert_eq!(
            resolve_color_value_to_rgba_hex(Some(&translucent), None).as_deref(),
            Some("#11223380")
        );
        assert_eq!(
            resolve_color_value_to_rgba_hex(
                Some(&ColorValue {
                    alpha: None,
                    ..translucent
                }),
                None
            )
            .as_deref(),
            Some("#112233FF")
        );
    }

    #[test]
    fn parses_and_resolves_direct_and_theme_colors() {
        let direct = parse_color_value(Some("AABBCC"), None, None, None);
        assert_eq!(
            resolve_color_value_to_hex(Some(&direct)).as_deref(),
            Some("#AABBCC")
        );

        let themed = parse_color_value(None, Some("accent1"), None, None);
        assert_eq!(
            resolve_color_value_to_hex(Some(&themed)).as_deref(),
            Some("#4472C4")
        );

        let mut theme = Theme::default();
        theme.color_scheme.accent1 = "204060".to_owned();
        let tinted = ColorValue {
            theme_color: Some("accent1".to_owned()),
            theme_tint: Some("80".to_owned()),
            ..ColorValue::default()
        };
        assert_eq!(
            resolve_color_value_to_hex_with_theme(Some(&tinted), Some(&theme)).as_deref(),
            Some("#90A0B0")
        );

        let malformed = ColorValue {
            rgb: Some("aéabc".to_owned()),
            ..ColorValue::default()
        };
        assert_eq!(resolve_color_value_to_hex(Some(&malformed)), None);
    }
}
