use ooxml_drawingml::{Theme, ThemeColorScheme, ThemeFont, ThemeFontScheme};

use crate::xml::XmlElement;

pub(crate) fn parse_theme(root: &XmlElement) -> Theme {
    let elements = root.child("themeElements");
    Theme {
        name: root.attribute("name").unwrap_or("Office Theme").to_owned(),
        color_scheme: parse_color_scheme(elements.and_then(|value| value.child("clrScheme"))),
        font_scheme: parse_font_scheme(elements.and_then(|value| value.child("fontScheme"))),
    }
}

fn parse_color_scheme(element: Option<&XmlElement>) -> ThemeColorScheme {
    let mut colors = ThemeColorScheme::default();
    let Some(element) = element else {
        return colors;
    };
    for slot in [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ] {
        let value = element
            .child(slot)
            .and_then(|slot| slot.child_elements().next())
            .and_then(parse_theme_color);
        if let Some(value) = value {
            colors.set(slot, value);
        }
    }
    colors
}

fn parse_theme_color(element: &XmlElement) -> Option<String> {
    match element.local_name() {
        "srgbClr" => element.attribute("val").map(str::to_owned),
        "sysClr" => element
            .attribute("lastClr")
            .map(str::to_owned)
            .or_else(|| system_color(element.attribute("val")).map(str::to_owned)),
        _ => None,
    }
}

fn system_color(value: Option<&str>) -> Option<&'static str> {
    match value? {
        "windowText" | "menuText" | "captionText" | "btnText" => Some("000000"),
        "window" | "menu" | "btnFace" | "btnHighlight" | "highlightText" => Some("FFFFFF"),
        "highlight" => Some("0078D7"),
        "grayText" => Some("808080"),
        _ => None,
    }
}

fn parse_font_scheme(element: Option<&XmlElement>) -> ThemeFontScheme {
    let Some(element) = element else {
        return ThemeFontScheme::default();
    };
    ThemeFontScheme {
        major_font: element
            .child("majorFont")
            .map(parse_theme_font)
            .unwrap_or_else(ThemeFont::default_major),
        minor_font: element
            .child("minorFont")
            .map(parse_theme_font)
            .unwrap_or_else(ThemeFont::default_minor),
    }
}

fn parse_theme_font(element: &XmlElement) -> ThemeFont {
    let mut font = ThemeFont::empty();
    font.latin = typeface(element.child("latin"));
    font.ea = typeface(element.child("ea"));
    font.cs = typeface(element.child("cs"));
    for script_font in element.children_named("font") {
        if let (Some(script), Some(typeface)) = (
            script_font.attribute("script"),
            script_font.attribute("typeface"),
        ) {
            font.fonts.insert(script.to_owned(), typeface.to_owned());
        }
    }
    font
}

fn typeface(element: Option<&XmlElement>) -> String {
    element
        .and_then(|value| value.attribute("typeface"))
        .unwrap_or_default()
        .to_owned()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    #[test]
    fn parses_presentation_theme_colors_and_fonts() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<a:theme name="BetterOffice"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="6254E7"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
            "ppt/theme/theme1.xml",
            &mut budget,
        )
        .unwrap();
        let theme = parse_theme(&root);
        assert_eq!(theme.name, "BetterOffice");
        assert_eq!(theme.color_scheme.accent1, "6254E7");
        assert_eq!(theme.font_scheme.minor_font.latin, "Aptos");
    }
}
