use serde::{Deserialize, Serialize};

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
    let color = color?;
    if let Some(rgb) = &color.rgb {
        return Some(format!("#{rgb}"));
    }
    color
        .theme_color
        .as_deref()
        .map(default_theme_color)
        .map(|rgb| format!("#{rgb}"))
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
    }
}
