//! DrawingML theme color/font subset required by S2 and future style resolution.

use crate::settings::ThemeFontLanguage;
use crate::settings::incumbent_utf8_text_boundary;
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

pub use ooxml_drawingml::{
    Theme, ThemeColorScheme, ThemeFont, ThemeFontScheme, get_default_theme, get_major_font,
    get_minor_font, get_theme_color, get_theme_fonts, resolve_theme_font_ref,
};

pub fn parse_theme(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Theme, ParseError> {
    let Some(xml) = xml.filter(|value| !is_blank(value)) else {
        return Ok(Theme::default());
    };
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(Theme::default());
    }
    let document = parse_xml(xml, part, budget)?;
    Ok(parse_theme_element(document.root()))
}

pub fn parse_theme_element(root: Option<&XmlElement>) -> Theme {
    let Some(root) = root else {
        return Theme::default();
    };
    let theme_elements = root.child("a", "themeElements");
    Theme {
        name: root
            .attribute(Some("a"), "name")
            .unwrap_or("Office Theme")
            .to_owned(),
        color_scheme: parse_color_scheme(
            theme_elements.and_then(|element| element.child("a", "clrScheme")),
        ),
        font_scheme: parse_font_scheme(
            theme_elements.and_then(|element| element.child("a", "fontScheme")),
        ),
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
            .child("a", slot)
            .and_then(|slot| slot.child_elements().next())
            .and_then(parse_theme_color_element);
        if let Some(value) = value {
            colors.set(slot, value);
        }
    }
    colors
}

fn parse_theme_color_element(element: &XmlElement) -> Option<String> {
    match element.local_name() {
        "srgbClr" => nonempty(element.attribute(Some("a"), "val")),
        "sysClr" => {
            if let Some(last_color) = nonempty(element.attribute(Some("a"), "lastClr")) {
                return Some(last_color);
            }
            match element.attribute(Some("a"), "val") {
                Some("windowText" | "menuText" | "captionText" | "btnText") => {
                    Some("000000".to_owned())
                }
                Some("window" | "menu" | "btnFace" | "btnHighlight" | "highlightText") => {
                    Some("FFFFFF".to_owned())
                }
                Some("highlight") => Some("0078D7".to_owned()),
                Some("grayText") => Some("808080".to_owned()),
                _ => None,
            }
        }
        // The incumbent deliberately leaves scheme references unresolved when
        // they occur inside the scheme itself.
        "schemeClr" => None,
        _ => None,
    }
}

fn parse_font_scheme(element: Option<&XmlElement>) -> ThemeFontScheme {
    let Some(element) = element else {
        return ThemeFontScheme::default();
    };
    ThemeFontScheme {
        major_font: element
            .child("a", "majorFont")
            .map(parse_theme_font)
            .unwrap_or_else(ThemeFont::default_major),
        minor_font: element
            .child("a", "minorFont")
            .map(parse_theme_font)
            .unwrap_or_else(ThemeFont::default_minor),
    }
}

fn parse_theme_font(element: &XmlElement) -> ThemeFont {
    let mut font = ThemeFont::empty();
    font.latin = typeface(element.child("a", "latin"));
    font.ea = typeface(element.child("a", "ea"));
    font.cs = typeface(element.child("a", "cs"));
    for script_font in element.children_named("a", "font") {
        let script = nonempty(script_font.attribute(Some("a"), "script"));
        let typeface = nonempty(script_font.attribute(Some("a"), "typeface"));
        if let (Some(script), Some(typeface)) = (script, typeface) {
            font.fonts.insert(script, typeface);
        }
    }
    font
}

fn typeface(element: Option<&XmlElement>) -> String {
    element
        .and_then(|element| element.attribute(Some("a"), "typeface"))
        .unwrap_or_default()
        .to_owned()
}

pub fn apply_theme_font_lang(theme: &mut Theme, language: Option<&ThemeFontLanguage>) {
    let Some(language) = language else { return };
    let east_asia_script = language
        .east_asia
        .as_deref()
        .and_then(east_asia_lang_to_script);
    let bidi_script = language.bidi.as_deref().and_then(bidi_lang_to_script);
    for font in [
        &mut theme.font_scheme.major_font,
        &mut theme.font_scheme.minor_font,
    ] {
        if font.ea.is_empty()
            && let Some(typeface) = east_asia_script.and_then(|script| font.fonts.get(script))
        {
            font.ea.clone_from(typeface);
        }
        if font.cs.is_empty()
            && let Some(typeface) = bidi_script.and_then(|script| font.fonts.get(script))
        {
            font.cs.clone_from(typeface);
        }
    }
}

fn east_asia_lang_to_script(language: &str) -> Option<&'static str> {
    let lower = language.to_ascii_lowercase();
    if lower.starts_with("ja") {
        Some("Jpan")
    } else if lower.starts_with("ko") {
        Some("Hang")
    } else if lower.starts_with("zh") {
        if lower.contains("hant")
            || lower
                .split('-')
                .any(|segment| matches!(segment, "tw" | "hk" | "mo"))
        {
            Some("Hant")
        } else {
            Some("Hans")
        }
    } else {
        None
    }
}

fn bidi_lang_to_script(language: &str) -> Option<&'static str> {
    let lower = language.to_ascii_lowercase();
    if lower.starts_with("ar") {
        Some("Arab")
    } else if lower.starts_with("he") || lower.starts_with("iw") {
        Some("Hebr")
    } else if lower.starts_with("th") {
        Some("Thai")
    } else if lower.starts_with("hi") || lower.starts_with("mr") || lower.starts_with("ne") {
        Some("Deva")
    } else {
        None
    }
}

fn nonempty(value: Option<&str>) -> Option<String> {
    value.filter(|value| !value.is_empty()).map(str::to_owned)
}

fn is_blank(value: &[u8]) -> bool {
    std::str::from_utf8(value).is_ok_and(|value| value.trim().is_empty())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn parse(xml: Option<&str>) -> Theme {
        let limits = ParseLimits::default();
        parse_theme(
            xml.map(str::as_bytes),
            "word/theme/theme1.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
    }

    #[test]
    fn defaults_for_missing_or_blank_input() {
        assert_eq!(parse(None), Theme::default());
        assert_eq!(parse(Some(" \n\t")), Theme::default());
    }

    #[test]
    fn parses_colors_fonts_system_fallbacks_and_duplicate_scripts() {
        let theme = parse(Some(
            r#"<a:theme name="Golden"><a:themeElements><a:clrScheme><a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1><a:accent1><a:srgbClr val="BADHEX"/></a:accent1><a:accent2><a:schemeClr val="phClr"/></a:accent2></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Major"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="First"/><a:font script="Jpan" typeface="Last"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
        ));
        assert_eq!(theme.name, "Golden");
        assert_eq!(theme.color_scheme.dk1, "000000");
        assert_eq!(theme.color_scheme.lt1, "FEFEFE");
        assert_eq!(theme.color_scheme.accent1, "BADHEX");
        assert_eq!(theme.color_scheme.accent2, "ED7D31");
        assert_eq!(theme.font_scheme.major_font.fonts["Jpan"], "Last");
    }

    #[test]
    fn resolves_language_slots_references_colors_and_unique_font_order() {
        let mut theme = parse(Some(
            r#"<a:theme><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Major"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hant" typeface="Traditional"/><a:font script="Arab" typeface="Arabic"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hant" typeface="Traditional"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
        ));
        apply_theme_font_lang(
            &mut theme,
            Some(&ThemeFontLanguage {
                east_asia: Some("zh-TW".to_owned()),
                bidi: Some("ar-SA".to_owned()),
            }),
        );
        assert_eq!(
            resolve_theme_font_ref(Some(&theme), "majorEastAsia"),
            "Traditional"
        );
        assert_eq!(resolve_theme_font_ref(Some(&theme), "minorAscii"), "Minor");
        assert_eq!(get_theme_color(Some(&theme), "missing"), "000000");
        assert_eq!(
            get_theme_fonts(Some(&theme)),
            ["Major", "Traditional", "Arabic", "Minor"]
        );
    }

    #[test]
    fn dtd_and_huge_attributes_are_rejected_without_panics() {
        let limits = ParseLimits::default();
        assert!(matches!(
            parse_theme(
                Some(b"<!DOCTYPE x><a:theme/>"),
                "word/theme/theme1.xml",
                &mut ParseBudget::new(&limits)
            ),
            Err(ParseError::UnsafeXml { .. })
        ));
        let huge = format!(r#"<a:theme name="{}"/>"#, "x".repeat(5_000_000));
        let limits = ParseLimits {
            max_attribute_bytes: 1024,
            ..ParseLimits::default()
        };
        assert!(matches!(
            parse_theme(
                Some(huge.as_bytes()),
                "word/theme/theme1.xml",
                &mut ParseBudget::new(&limits)
            ),
            Err(ParseError::ResourceLimit { .. })
        ));
    }
}
