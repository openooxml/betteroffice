use serde::{Deserialize, Serialize};

use crate::Theme;

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
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
    Some(format!(
        "#{:02X}{:02X}{:02X}",
        channels[0], channels[1], channels[2]
    ))
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
