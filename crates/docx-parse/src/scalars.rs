//! Shared scalar property decoders used by borders and the later style/run slices.

use serde::{Deserialize, Serialize};

use crate::xml::XmlElement;

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

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShadingProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pattern: Option<String>,
}

pub fn parse_shading_properties(shading: Option<&XmlElement>) -> Option<ShadingProperties> {
    let shading = shading?;
    let mut properties = ShadingProperties::default();

    let color = shading.attribute(Some("w"), "color");
    if let Some(color) = color.filter(|value| !value.is_empty() && *value != "auto") {
        properties.color = Some(ColorValue {
            rgb: Some(color.to_owned()),
            ..ColorValue::default()
        });
    }
    if let Some(theme_color) = shading
        .attribute(Some("w"), "themeColor")
        .filter(|value| !value.is_empty())
    {
        let value = properties.color.get_or_insert_with(ColorValue::default);
        value.theme_color = Some(theme_color.to_owned());
        value.theme_tint = shading.attribute(Some("w"), "themeTint").map(str::to_owned);
        value.theme_shade = shading
            .attribute(Some("w"), "themeShade")
            .map(str::to_owned);
    }

    let fill = shading.attribute(Some("w"), "fill");
    if let Some(fill) = fill.filter(|value| !value.is_empty() && *value != "auto") {
        properties.fill = Some(ColorValue {
            rgb: Some(fill.to_owned()),
            ..ColorValue::default()
        });
    }
    if let Some(theme_fill) = shading
        .attribute(Some("w"), "themeFill")
        .filter(|value| !value.is_empty())
    {
        properties
            .fill
            .get_or_insert_with(ColorValue::default)
            .theme_color = Some(theme_fill.to_owned());
    }
    if let (Some(theme_tint), Some(value)) = (
        shading
            .attribute(Some("w"), "themeFillTint")
            .filter(|value| !value.is_empty()),
        properties.fill.as_mut(),
    ) {
        value.theme_tint = Some(theme_tint.to_owned());
    }
    if let (Some(theme_shade), Some(value)) = (
        shading
            .attribute(Some("w"), "themeFillShade")
            .filter(|value| !value.is_empty()),
        properties.fill.as_mut(),
    ) {
        value.theme_shade = Some(theme_shade.to_owned());
    }
    properties.pattern = shading
        .attribute(Some("w"), "val")
        .filter(|value| !value.is_empty())
        .map(str::to_owned);

    (properties != ShadingProperties::default()).then_some(properties)
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct UnderlineValue {
    pub style: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
}

pub fn parse_underline(element: Option<&XmlElement>) -> Option<UnderlineValue> {
    let element = element?;
    let style = element.attribute(Some("w"), "val")?;
    if style.is_empty() {
        return None;
    }
    let rgb = element.attribute(Some("w"), "color");
    let theme_color = element.attribute(Some("w"), "themeColor");
    let color = (rgb.is_some() || theme_color.is_some()).then(|| {
        parse_color_value(
            rgb,
            theme_color,
            element.attribute(Some("w"), "themeTint"),
            element.attribute(Some("w"), "themeShade"),
        )
    });
    Some(UnderlineValue {
        style: style.to_owned(),
        color,
    })
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunScalarProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<ShadingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kerning: Option<f64>,
}

pub fn parse_run_scalar_properties(r_pr: Option<&XmlElement>) -> Option<RunScalarProperties> {
    let r_pr = r_pr?;
    let color = r_pr.child("w", "color").map(|element| {
        parse_color_value(
            element.attribute(Some("w"), "val"),
            element.attribute(Some("w"), "themeColor"),
            element.attribute(Some("w"), "themeTint"),
            element.attribute(Some("w"), "themeShade"),
        )
    });
    let numeric = |name: &str| {
        r_pr.child("w", name)
            .and_then(|element| element.parse_numeric_attribute(Some("w"), "val", 1.0))
    };
    let properties = RunScalarProperties {
        underline: parse_underline(r_pr.child("w", "u")),
        color,
        highlight: r_pr
            .child("w", "highlight")
            .and_then(|element| element.attribute(Some("w"), "val"))
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        shading: parse_shading_properties(r_pr.child("w", "shd")),
        font_size: numeric("sz"),
        font_size_cs: numeric("szCs"),
        spacing: numeric("spacing"),
        position: numeric("position"),
        scale: numeric("w"),
        kerning: numeric("kern"),
    };
    (properties != RunScalarProperties::default()).then_some(properties)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(xml.as_bytes(), "scalar.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn decodes_color_shading_underline_highlight_and_measurements() {
        let r_pr = root(
            r#"<w:rPr><w:u w:val="double" w:color="auto"/><w:color w:val="bad-hex" w:themeColor="accent1"/><w:highlight w:val="yellow"/><w:shd w:fill="auto" w:themeFill="accent2" w:themeFillTint="80" w:val="clear"/><w:sz w:val="24half"/><w:spacing w:val="-20twips"/></w:rPr>"#,
        );
        let parsed = parse_run_scalar_properties(Some(&r_pr)).unwrap();
        assert_eq!(parsed.underline.unwrap().color.unwrap().auto, Some(true));
        assert_eq!(parsed.color.unwrap().rgb.as_deref(), Some("bad-hex"));
        assert_eq!(parsed.highlight.as_deref(), Some("yellow"));
        assert_eq!(
            parsed.shading.unwrap().fill.unwrap().theme_tint.as_deref(),
            Some("80")
        );
        assert_eq!(parsed.font_size, Some(24.0));
        assert_eq!(parsed.spacing, Some(-20.0));
    }

    #[test]
    fn malformed_and_huge_numeric_values_never_become_non_finite() {
        for value in ["", "garbage", "+", &"9".repeat(10_000)] {
            let r_pr = root(&format!("<w:rPr><w:sz w:val=\"{value}\"/></w:rPr>"));
            let parsed = parse_run_scalar_properties(Some(&r_pr));
            assert!(parsed.and_then(|value| value.font_size).is_none());
        }
        let prefixed = root(r#"<w:rPr><w:sz w:val="9e999"/></w:rPr>"#);
        assert_eq!(
            parse_run_scalar_properties(Some(&prefixed)).and_then(|value| value.font_size),
            Some(9.0)
        );
    }
}
