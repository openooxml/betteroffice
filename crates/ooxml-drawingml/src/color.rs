use serde::{Deserialize, Serialize};

use crate::Theme;

/// DrawingML transforms; hue and hue offsets use turns, other values use fractions.
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", content = "value", rename_all = "camelCase")]
pub enum ColorTransform {
    Tint(f64),
    Shade(f64),
    Hue(f64),
    HueMod(f64),
    HueOff(f64),
    Sat(f64),
    SatMod(f64),
    SatOff(f64),
    Lum(f64),
    LumMod(f64),
    LumOff(f64),
    Red(f64),
    RedMod(f64),
    RedOff(f64),
    Green(f64),
    GreenMod(f64),
    GreenOff(f64),
    Blue(f64),
    BlueMod(f64),
    BlueOff(f64),
    Alpha(f64),
    AlphaMod(f64),
    AlphaOff(f64),
    Comp,
    Inv,
    Gray,
    Gamma,
    InvGamma,
}

impl ColorTransform {
    pub fn from_xml(name: &str, value: Option<&str>) -> Option<Self> {
        match name {
            "comp" => return Some(Self::Comp),
            "inv" => return Some(Self::Inv),
            "gray" => return Some(Self::Gray),
            "gamma" => return Some(Self::Gamma),
            "invGamma" => return Some(Self::InvGamma),
            _ => {}
        }
        let raw = value?.trim();
        let value = if matches!(name, "hue" | "hueOff") {
            raw.parse::<f64>().ok()? / 21_600_000.0
        } else if let Some(percent) = raw.strip_suffix('%') {
            percent.parse::<f64>().ok()? / 100.0
        } else {
            raw.parse::<f64>().ok()? / 100_000.0
        };
        if !value.is_finite() {
            return None;
        }
        let in_range = match name {
            "tint" | "shade" => (0.0..=1.0).contains(&value),
            "lumMod" | "lumOff" | "satMod" | "alpha" => (0.0..=10.0).contains(&value),
            _ => true,
        };
        if !in_range {
            return None;
        }
        Some(match name {
            "tint" => Self::Tint(value),
            "shade" => Self::Shade(value),
            "hue" => Self::Hue(value),
            "hueMod" => Self::HueMod(value),
            "hueOff" => Self::HueOff(value),
            "sat" => Self::Sat(value),
            "satMod" => Self::SatMod(value),
            "satOff" => Self::SatOff(value),
            "lum" => Self::Lum(value),
            "lumMod" => Self::LumMod(value),
            "lumOff" => Self::LumOff(value),
            "red" => Self::Red(value),
            "redMod" => Self::RedMod(value),
            "redOff" => Self::RedOff(value),
            "green" => Self::Green(value),
            "greenMod" => Self::GreenMod(value),
            "greenOff" => Self::GreenOff(value),
            "blue" => Self::Blue(value),
            "blueMod" => Self::BlueMod(value),
            "blueOff" => Self::BlueOff(value),
            "alpha" => Self::Alpha(value),
            "alphaMod" => Self::AlphaMod(value),
            "alphaOff" => Self::AlphaOff(value),
            _ => return None,
        })
    }

    pub fn xml_name(&self) -> &'static str {
        match self {
            Self::Tint(_) => "tint",
            Self::Shade(_) => "shade",
            Self::Hue(_) => "hue",
            Self::HueMod(_) => "hueMod",
            Self::HueOff(_) => "hueOff",
            Self::Sat(_) => "sat",
            Self::SatMod(_) => "satMod",
            Self::SatOff(_) => "satOff",
            Self::Lum(_) => "lum",
            Self::LumMod(_) => "lumMod",
            Self::LumOff(_) => "lumOff",
            Self::Red(_) => "red",
            Self::RedMod(_) => "redMod",
            Self::RedOff(_) => "redOff",
            Self::Green(_) => "green",
            Self::GreenMod(_) => "greenMod",
            Self::GreenOff(_) => "greenOff",
            Self::Blue(_) => "blue",
            Self::BlueMod(_) => "blueMod",
            Self::BlueOff(_) => "blueOff",
            Self::Alpha(_) => "alpha",
            Self::AlphaMod(_) => "alphaMod",
            Self::AlphaOff(_) => "alphaOff",
            Self::Comp => "comp",
            Self::Inv => "inv",
            Self::Gray => "gray",
            Self::Gamma => "gamma",
            Self::InvGamma => "invGamma",
        }
    }

    pub fn xml_value(&self) -> Option<String> {
        let value = match *self {
            Self::Hue(value)
            | Self::HueOff(value)
            | Self::Tint(value)
            | Self::Shade(value)
            | Self::HueMod(value)
            | Self::Sat(value)
            | Self::SatMod(value)
            | Self::SatOff(value)
            | Self::Lum(value)
            | Self::LumMod(value)
            | Self::LumOff(value)
            | Self::Red(value)
            | Self::RedMod(value)
            | Self::RedOff(value)
            | Self::Green(value)
            | Self::GreenMod(value)
            | Self::GreenOff(value)
            | Self::Blue(value)
            | Self::BlueMod(value)
            | Self::BlueOff(value)
            | Self::Alpha(value)
            | Self::AlphaMod(value)
            | Self::AlphaOff(value) => value,
            _ => return None,
        };
        let angle = matches!(self, Self::Hue(_) | Self::HueOff(_));
        let scaled = value * if angle { 21_600_000.0 } else { 100_000.0 };
        if !scaled.is_finite() {
            return None;
        }
        if angle || (scaled - scaled.round()).abs() < 1e-8 {
            Some(scaled.round().to_string())
        } else {
            Some(format!("{}%", value * 100.0))
        }
    }
}

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
    /// Overrides legacy colour modifiers; legacy alpha applies without an alpha step.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub transforms: Vec<ColorTransform>,
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

impl ColorValue {
    /// Ordered transforms with legacy alpha as a fallback.
    pub fn drawingml_transforms(&self) -> impl Iterator<Item = ColorTransform> + '_ {
        let alpha = (!self.transforms.is_empty()
            && !self.transforms.iter().any(|transform| {
                matches!(
                    transform,
                    ColorTransform::Alpha(_)
                        | ColorTransform::AlphaMod(_)
                        | ColorTransform::AlphaOff(_)
                )
            }))
        .then_some(self.alpha)
        .flatten()
        .filter(|alpha| alpha.is_finite())
        .map(|alpha| ColorTransform::Alpha(alpha.clamp(0.0, 1.0)));
        self.transforms.iter().copied().chain(alpha)
    }

    /// A plain `w:val`/`w:fill` attribute, where `auto` defers to the consumer.
    pub fn from_attribute(value: &str) -> Self {
        parse_color_value(Some(value), None, None, None)
    }
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
                .and_then(|theme| theme.color_scheme.get(theme.color_map.resolve(slot)))
                .unwrap_or_else(|| default_theme_color(slot))
        })
    })?;
    let mut channels = parse_rgb(rgb)?;
    if !color.transforms.is_empty() {
        let channels = apply_transforms(
            channels.map(|channel| f64::from(channel) / 255.0),
            &color.transforms,
        )
        .map(|channel| (channel * 255.0).round() as u8);
        return Some(format!(
            "#{:02X}{:02X}{:02X}",
            channels[0], channels[1], channels[2]
        ));
    }
    if let Some(shade) = color.theme_shade.as_deref().and_then(parse_modifier) {
        channels = channels.map(|channel| (f64::from(channel) * shade).round() as u8);
    }
    if let Some(tint) = color.theme_tint.as_deref().and_then(parse_modifier) {
        channels = channels
            .map(|channel| (f64::from(channel) * tint + 255.0 * (1.0 - tint)).round() as u8);
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
    let alpha = color.map_or(1.0, resolve_color_value_alpha);
    Some(format!("{hex}{:02X}", (alpha * 255.0).round() as u8))
}

pub fn resolve_color_value_alpha(color: &ColorValue) -> f64 {
    if color.transforms.is_empty() {
        return color
            .alpha
            .filter(|alpha| alpha.is_finite() && *alpha < 1.0)
            .unwrap_or(1.0)
            .clamp(0.0, 1.0);
    }
    let mut alpha = 1.0;
    for transform in color.drawingml_transforms() {
        alpha = match transform {
            ColorTransform::Alpha(value) if value.is_finite() => value,
            ColorTransform::AlphaMod(value) if value.is_finite() => alpha * value,
            ColorTransform::AlphaOff(value) if value.is_finite() => alpha + value,
            _ => alpha,
        }
        .clamp(0.0, 1.0);
    }
    alpha
}

fn srgb_to_linear(value: f64) -> f64 {
    if value <= 0.04045 {
        value / 12.92
    } else {
        ((value + 0.055) / 1.055).powf(2.4)
    }
}

fn linear_to_srgb(value: f64) -> f64 {
    if value <= 0.0031308 {
        value * 12.92
    } else {
        1.055 * value.powf(1.0 / 2.4) - 0.055
    }
}

fn apply_transforms(mut channels: [f64; 3], transforms: &[ColorTransform]) -> [f64; 3] {
    use ColorTransform::*;

    for transform in transforms {
        match *transform {
            Tint(value) | Shade(value) if value.is_finite() => {
                channels = channels.map(|channel| {
                    let linear = srgb_to_linear(channel) * value;
                    let linear = if matches!(transform, Tint(_)) {
                        linear + (1.0 - value)
                    } else {
                        linear
                    };
                    linear_to_srgb(linear.clamp(0.0, 1.0))
                });
            }
            Hue(value) | HueMod(value) | HueOff(value) | Sat(value) | SatMod(value)
            | SatOff(value) | Lum(value) | LumMod(value) | LumOff(value)
                if value.is_finite() =>
            {
                let (mut hue, mut sat, mut lum) = rgb_to_hsl_f64(channels);
                match *transform {
                    Hue(_) => hue = value * 360.0,
                    HueMod(_) => hue *= value,
                    HueOff(_) => hue += value * 360.0,
                    Sat(_) => sat = value,
                    SatMod(_) => sat *= value,
                    SatOff(_) => sat += value,
                    Lum(_) => lum = value,
                    LumMod(_) => lum *= value,
                    LumOff(_) => lum += value,
                    _ => unreachable!(),
                }
                channels = hsl_to_rgb_f64(
                    hue.rem_euclid(360.0),
                    sat.clamp(0.0, 1.0),
                    lum.clamp(0.0, 1.0),
                );
            }
            Comp => {
                let (hue, sat, lum) = rgb_to_hsl_f64(channels);
                channels = hsl_to_rgb_f64(hue + 180.0, sat, lum);
            }
            Red(value) | RedMod(value) | RedOff(value) | Green(value) | GreenMod(value)
            | GreenOff(value) | Blue(value) | BlueMod(value) | BlueOff(value)
                if value.is_finite() =>
            {
                let index = match transform {
                    Red(_) | RedMod(_) | RedOff(_) => 0,
                    Green(_) | GreenMod(_) | GreenOff(_) => 1,
                    _ => 2,
                };
                let linear = srgb_to_linear(channels[index]);
                let linear = match transform {
                    RedMod(_) | GreenMod(_) | BlueMod(_) => linear * value,
                    RedOff(_) | GreenOff(_) | BlueOff(_) => linear + value,
                    _ => value,
                };
                channels[index] = linear_to_srgb(linear.clamp(0.0, 1.0));
            }
            Inv => channels = channels.map(|channel| linear_to_srgb(1.0 - srgb_to_linear(channel))),
            Gray => {
                let [red, green, blue] = channels.map(srgb_to_linear);
                channels = [linear_to_srgb(0.2126 * red + 0.7152 * green + 0.0722 * blue); 3];
            }
            Gamma => channels = channels.map(linear_to_srgb),
            InvGamma => channels = channels.map(srgb_to_linear),
            _ => {}
        }
        channels = channels.map(|channel| channel.clamp(0.0, 1.0));
    }
    channels
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
    rgb_to_hsl_f64(channels.map(|channel| f64::from(channel) / 255.0))
}

fn rgb_to_hsl_f64([red, green, blue]: [f64; 3]) -> (f64, f64, f64) {
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
    hsl_to_rgb_f64(hue, sat, lum).map(|channel| (channel * 255.0).round() as u8)
}

fn hsl_to_rgb_f64(hue: f64, sat: f64, lum: f64) -> [f64; 3] {
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
    [red, green, blue].map(|channel| (channel + base).clamp(0.0, 1.0))
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
    fn transform_xml_values_reject_non_finite_numbers() {
        use ColorTransform::*;

        let constructors: [fn(f64) -> ColorTransform; 23] = [
            Tint, Shade, Hue, HueMod, HueOff, Sat, SatMod, SatOff, Lum, LumMod, LumOff, Red,
            RedMod, RedOff, Green, GreenMod, GreenOff, Blue, BlueMod, BlueOff, Alpha, AlphaMod,
            AlphaOff,
        ];
        for constructor in constructors {
            for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, f64::MAX] {
                assert_eq!(constructor(value).xml_value(), None);
            }
            assert!(constructor(0.5).xml_value().is_some());
        }
    }

    #[test]
    fn transform_parsing_retains_legacy_ranges() {
        for (name, maximum) in [
            ("tint", 1.0),
            ("shade", 1.0),
            ("lumMod", 10.0),
            ("lumOff", 10.0),
            ("satMod", 10.0),
            ("alpha", 10.0),
        ] {
            for value in [-0.00001, maximum + 0.00001] {
                let raw = (value * 100_000.0).to_string();
                assert_eq!(ColorTransform::from_xml(name, Some(&raw)), None, "{name}");
            }
            for value in [0.0, maximum] {
                for raw in [
                    (value * 100_000.0).to_string(),
                    format!("{}%", value * 100.0),
                ] {
                    assert!(
                        ColorTransform::from_xml(name, Some(&raw)).is_some(),
                        "{name}"
                    );
                }
            }
            for raw in ["NaN", "inf", "-inf"] {
                assert_eq!(ColorTransform::from_xml(name, Some(raw)), None);
            }
        }
        assert_eq!(
            ColorTransform::from_xml("satOff", Some("-50%")),
            Some(ColorTransform::SatOff(-0.5))
        );
        assert_eq!(
            ColorTransform::from_xml("hueOff", Some("-5400000")),
            Some(ColorTransform::HueOff(-0.25))
        );
        assert_eq!(
            ColorTransform::from_xml("alphaMod", Some("200000")),
            Some(ColorTransform::AlphaMod(2.0))
        );
    }

    #[test]
    fn mixed_color_keeps_legacy_alpha_without_an_alpha_step() {
        let mut color = ColorValue {
            rgb: Some("112233".to_owned()),
            theme_tint: Some("00".to_owned()),
            luminance_modulation: Some(0.0),
            alpha: Some(0.25),
            transforms: vec![ColorTransform::Tint(1.0)],
            ..Default::default()
        };
        assert_eq!(
            resolve_color_value_to_rgba_hex(Some(&color), None).as_deref(),
            Some("#11223340")
        );
        for (step, expected) in [
            (ColorTransform::Alpha(0.5), 0.5),
            (ColorTransform::AlphaMod(0.5), 0.5),
            (ColorTransform::AlphaOff(-0.25), 0.75),
        ] {
            color.transforms = vec![step];
            assert_eq!(resolve_color_value_alpha(&color), expected);
        }
    }

    #[test]
    fn transforms_use_their_own_color_spaces() {
        use ColorTransform::*;

        for (rgb, transforms, expected) in [
            ("000000", vec![Red(0.5)], "#BC0000"),
            ("000000", vec![Green(0.5)], "#00BC00"),
            ("000000", vec![Blue(0.5)], "#0000BC"),
            ("FFFFFF", vec![RedMod(0.5)], "#BCFFFF"),
            ("FFFFFF", vec![GreenMod(0.5)], "#FFBCFF"),
            ("FFFFFF", vec![BlueMod(0.5)], "#FFFFBC"),
            ("000000", vec![RedOff(0.5)], "#BC0000"),
            ("000000", vec![GreenOff(0.5)], "#00BC00"),
            ("000000", vec![BlueOff(0.5)], "#0000BC"),
            ("808080", vec![Inv], "#E5E5E5"),
            ("FF0000", vec![Gray], "#7F7F7F"),
            ("808080", vec![Gamma], "#BCBCBC"),
            ("808080", vec![InvGamma], "#373737"),
            ("808080", vec![Gamma, InvGamma], "#808080"),
            ("FF0000", vec![Hue(1.0 / 3.0)], "#00FF00"),
            ("00FF00", vec![HueMod(2.0)], "#0000FF"),
            ("FF0000", vec![HueOff(-1.0 / 3.0)], "#0000FF"),
            ("FF0000", vec![Comp], "#00FFFF"),
            ("FF0000", vec![Sat(0.0)], "#808080"),
            ("FF0000", vec![SatMod(0.5)], "#BF4040"),
            ("FF0000", vec![SatOff(-0.5)], "#BF4040"),
            ("FF0000", vec![Lum(0.25)], "#800000"),
            ("FF0000", vec![LumOff(0.25)], "#FF8080"),
            ("FF0000", vec![LumOff(1.0), LumMod(0.5)], "#808080"),
            ("000000", vec![RedOff(-0.5), RedOff(0.5)], "#BC0000"),
        ] {
            let color = ColorValue {
                rgb: Some(rgb.to_owned()),
                transforms,
                ..Default::default()
            };
            assert_eq!(
                resolve_color_value_to_hex(Some(&color)).as_deref(),
                Some(expected),
                "{color:?}"
            );
        }
    }

    #[test]
    fn ordered_alpha_is_separate_and_overrides_legacy_fields() {
        let mut color = ColorValue {
            rgb: Some("112233".to_owned()),
            theme_tint: Some("00".to_owned()),
            luminance_modulation: Some(0.0),
            alpha: Some(0.0),
            transforms: vec![
                ColorTransform::Alpha(0.5),
                ColorTransform::AlphaMod(0.5),
                ColorTransform::AlphaOff(0.25),
            ],
            ..Default::default()
        };
        assert_eq!(
            resolve_color_value_to_rgba_hex(Some(&color), None).as_deref(),
            Some("#11223380")
        );
        assert_eq!(
            resolve_color_value_to_hex(Some(&color)).as_deref(),
            Some("#112233")
        );
        color.transforms = vec![
            ColorTransform::AlphaOff(-2.0),
            ColorTransform::AlphaOff(0.5),
        ];
        assert_eq!(resolve_color_value_alpha(&color), 0.5);
        color.transforms = vec![ColorTransform::Alpha(0.25), ColorTransform::Alpha(1.0)];
        assert_eq!(resolve_color_value_alpha(&color), 1.0);
    }

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
    fn legacy_tint_and_shade_keep_their_byte_math() {
        let modified = |tint: Option<&str>, shade: Option<&str>| {
            resolve_color_value_to_hex(Some(&ColorValue {
                rgb: Some("204060".to_owned()),
                theme_tint: tint.map(str::to_owned),
                theme_shade: shade.map(str::to_owned),
                ..ColorValue::default()
            }))
            .expect("resolves")
        };

        assert_eq!(modified(Some("FF"), None), "#204060");
        assert_eq!(modified(Some("CC"), None), "#4D6680");
        assert_eq!(modified(Some("80"), None), "#8F9FAF");
        assert_eq!(modified(Some("33"), None), "#D2D9DF");
        assert_eq!(modified(Some("00"), None), "#FFFFFF");

        assert_eq!(modified(None, Some("FF")), "#204060");
        assert_eq!(modified(None, Some("80")), "#102030");
        assert_eq!(modified(None, Some("33")), "#060D13");
        assert_eq!(modified(None, Some("00")), "#000000");

        const LEGACY_ACCENT1_TINT_66: &str = "#B4C7E7";
        assert_eq!(
            resolve_color_value_to_hex(Some(&ColorValue {
                theme_color: Some("accent1".to_owned()),
                theme_tint: Some("66".to_owned()),
                ..ColorValue::default()
            }))
            .as_deref(),
            Some(LEGACY_ACCENT1_TINT_66)
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
            Some("#8F9FAF")
        );

        let malformed = ColorValue {
            rgb: Some("aéabc".to_owned()),
            ..ColorValue::default()
        };
        assert_eq!(resolve_color_value_to_hex(Some(&malformed)), None);
    }
}
