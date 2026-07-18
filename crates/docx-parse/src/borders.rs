//! Shared `CT_Border` parsing for table, paragraph, and page consumers.

use serde::{Deserialize, Serialize};

use crate::scalars::{ColorValue, parse_color_value};
use crate::xml::XmlElement;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderSpec {
    pub style: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shadow: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frame: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Borders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inside_h: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inside_v: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub between: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tl2br: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tr2bl: Option<BorderSpec>,
}

pub fn parse_border_spec(element: Option<&XmlElement>) -> Option<BorderSpec> {
    let element = element?;
    let color_value = element
        .attribute(Some("w"), "color")
        .filter(|value| !value.is_empty());
    let theme_color = element
        .attribute(Some("w"), "themeColor")
        .filter(|value| !value.is_empty());
    let theme_tint = element
        .attribute(Some("w"), "themeTint")
        .filter(|value| !value.is_empty());
    let theme_shade = element
        .attribute(Some("w"), "themeShade")
        .filter(|value| !value.is_empty());
    Some(BorderSpec {
        style: element
            .attribute(Some("w"), "val")
            .unwrap_or("none")
            .to_owned(),
        color: (color_value.is_some()
            || theme_color.is_some()
            || theme_tint.is_some()
            || theme_shade.is_some())
        .then(|| parse_color_value(color_value, theme_color, theme_tint, theme_shade)),
        size: element.parse_numeric_attribute(Some("w"), "sz", 1.0),
        space: element.parse_numeric_attribute(Some("w"), "space", 1.0),
        shadow: matches!(element.attribute(Some("w"), "shadow"), Some("1" | "true"))
            .then_some(true),
        frame: matches!(element.attribute(Some("w"), "frame"), Some("1" | "true")).then_some(true),
    })
}

pub fn parse_table_borders(element: Option<&XmlElement>) -> Option<Borders> {
    let element = element?;
    let borders = Borders {
        top: parse_border_spec(element.child("w", "top")),
        bottom: parse_border_spec(element.child("w", "bottom")),
        left: parse_border_spec(
            element
                .child("w", "left")
                .or_else(|| element.child("w", "start")),
        ),
        right: parse_border_spec(
            element
                .child("w", "right")
                .or_else(|| element.child("w", "end")),
        ),
        inside_h: parse_border_spec(element.child("w", "insideH")),
        inside_v: parse_border_spec(element.child("w", "insideV")),
        ..Borders::default()
    };
    (borders != Borders::default()).then_some(borders)
}

pub fn parse_paragraph_borders(element: Option<&XmlElement>) -> Option<Borders> {
    let element = element?;
    let borders = Borders {
        top: parse_border_spec(element.child("w", "top")),
        bottom: parse_border_spec(element.child("w", "bottom")),
        left: parse_border_spec(element.child("w", "left")),
        right: parse_border_spec(element.child("w", "right")),
        between: parse_border_spec(element.child("w", "between")),
        bar: parse_border_spec(element.child("w", "bar")),
        ..Borders::default()
    };
    (borders != Borders::default()).then_some(borders)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(xml.as_bytes(), "border.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn preserves_nil_auto_theme_and_rtl_alias_behavior() {
        let borders = root(
            r#"<w:tblBorders><w:top w:val="nil"/><w:start w:val="single" w:sz="4px" w:color="auto" w:themeTint="80"/><w:end/></w:tblBorders>"#,
        );
        let parsed = parse_table_borders(Some(&borders)).unwrap();
        assert_eq!(parsed.top.unwrap().style, "nil");
        let left = parsed.left.unwrap();
        assert_eq!(left.size, Some(4.0));
        assert_eq!(left.color.unwrap().auto, Some(true));
        assert_eq!(parsed.right.unwrap().style, "none");
    }

    #[test]
    fn malformed_numbers_are_omitted_without_panics() {
        let border = root(&format!(
            r#"<w:top w:val="single" w:sz="{}" w:space="garbage"/>"#,
            "9".repeat(10_000)
        ));
        let parsed = parse_border_spec(Some(&border)).unwrap();
        assert_eq!(parsed.size, None);
        assert_eq!(parsed.space, None);
    }
}
