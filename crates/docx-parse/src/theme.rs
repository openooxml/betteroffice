//! DrawingML theme color/font subset required by S2 and future style resolution.

use indexmap::{IndexMap, IndexSet};
use serde::{Deserialize, Serialize};

use crate::settings::ThemeFontLanguage;
use crate::settings::incumbent_utf8_text_boundary;
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeColorScheme {
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub fol_hlink: String,
}

impl Default for ThemeColorScheme {
    fn default() -> Self {
        Self {
            dk1: "000000".to_owned(),
            lt1: "FFFFFF".to_owned(),
            dk2: "44546A".to_owned(),
            lt2: "E7E6E6".to_owned(),
            accent1: "4472C4".to_owned(),
            accent2: "ED7D31".to_owned(),
            accent3: "A5A5A5".to_owned(),
            accent4: "FFC000".to_owned(),
            accent5: "5B9BD5".to_owned(),
            accent6: "70AD47".to_owned(),
            hlink: "0563C1".to_owned(),
            fol_hlink: "954F72".to_owned(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeFont {
    pub latin: String,
    pub ea: String,
    pub cs: String,
    pub fonts: IndexMap<String, String>,
}

impl ThemeFont {
    fn default_major() -> Self {
        Self {
            latin: "Calibri Light".to_owned(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }

    fn default_minor() -> Self {
        Self {
            latin: "Calibri".to_owned(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }

    fn empty() -> Self {
        Self {
            latin: String::new(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFontScheme {
    pub major_font: ThemeFont,
    pub minor_font: ThemeFont,
}

impl Default for ThemeFontScheme {
    fn default() -> Self {
        Self {
            major_font: ThemeFont::default_major(),
            minor_font: ThemeFont::default_minor(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Theme {
    pub name: String,
    pub color_scheme: ThemeColorScheme,
    pub font_scheme: ThemeFontScheme,
}

impl Default for Theme {
    fn default() -> Self {
        Self {
            name: "Office Theme".to_owned(),
            color_scheme: ThemeColorScheme::default(),
            font_scheme: ThemeFontScheme::default(),
        }
    }
}

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

impl ThemeColorScheme {
    fn set(&mut self, slot: &str, value: String) {
        match slot {
            "dk1" => self.dk1 = value,
            "lt1" => self.lt1 = value,
            "dk2" => self.dk2 = value,
            "lt2" => self.lt2 = value,
            "accent1" => self.accent1 = value,
            "accent2" => self.accent2 = value,
            "accent3" => self.accent3 = value,
            "accent4" => self.accent4 = value,
            "accent5" => self.accent5 = value,
            "accent6" => self.accent6 = value,
            "hlink" => self.hlink = value,
            "folHlink" => self.fol_hlink = value,
            _ => {}
        }
    }

    fn get(&self, slot: &str) -> Option<&str> {
        match slot {
            "dk1" => Some(&self.dk1),
            "lt1" => Some(&self.lt1),
            "dk2" => Some(&self.dk2),
            "lt2" => Some(&self.lt2),
            "accent1" => Some(&self.accent1),
            "accent2" => Some(&self.accent2),
            "accent3" => Some(&self.accent3),
            "accent4" => Some(&self.accent4),
            "accent5" => Some(&self.accent5),
            "accent6" => Some(&self.accent6),
            "hlink" => Some(&self.hlink),
            "folHlink" => Some(&self.fol_hlink),
            _ => None,
        }
        .map(String::as_str)
    }
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

pub fn get_theme_color(theme: Option<&Theme>, slot: &str) -> String {
    if let Some(value) = theme.and_then(|theme| theme.color_scheme.get(slot)) {
        return value.to_owned();
    }
    let defaults = ThemeColorScheme::default();
    defaults.get(slot).unwrap_or("000000").to_owned()
}

pub fn get_major_font(theme: Option<&Theme>, script: &str) -> String {
    get_font(
        theme.map(|theme| &theme.font_scheme.major_font),
        script,
        "Calibri Light",
    )
}

pub fn get_minor_font(theme: Option<&Theme>, script: &str) -> String {
    get_font(
        theme.map(|theme| &theme.font_scheme.minor_font),
        script,
        "Calibri",
    )
}

fn get_font(font: Option<&ThemeFont>, script: &str, latin_default: &str) -> String {
    let Some(font) = font else {
        return latin_default.to_owned();
    };
    match script {
        "latin" => nonempty(Some(&font.latin)).unwrap_or_else(|| latin_default.to_owned()),
        "ea" => font.ea.clone(),
        "cs" => font.cs.clone(),
        script => font
            .fonts
            .get(script)
            .cloned()
            .or_else(|| nonempty(Some(&font.latin)))
            .unwrap_or_else(|| latin_default.to_owned()),
    }
}

pub fn resolve_theme_font_ref(theme: Option<&Theme>, reference: &str) -> String {
    if reference.is_empty() {
        return "Calibri".to_owned();
    }
    let lower = reference.to_ascii_lowercase();
    let script = if lower.contains("eastasia") {
        "ea"
    } else if lower.contains("bidi") || lower.contains("cs") {
        "cs"
    } else {
        "latin"
    };
    if lower.contains("major") {
        get_major_font(theme, script)
    } else {
        get_minor_font(theme, script)
    }
}

pub fn get_theme_fonts(theme: Option<&Theme>) -> Vec<String> {
    let mut fonts = IndexSet::new();
    if let Some(theme) = theme {
        for font in [&theme.font_scheme.major_font, &theme.font_scheme.minor_font] {
            for value in [&font.latin, &font.ea, &font.cs] {
                if !value.is_empty() {
                    fonts.insert(value.clone());
                }
            }
        }
        for value in theme
            .font_scheme
            .major_font
            .fonts
            .values()
            .chain(theme.font_scheme.minor_font.fonts.values())
        {
            if !value.is_empty() {
                fonts.insert(value.clone());
            }
        }
    }
    fonts.into_iter().collect()
}

pub fn get_default_theme() -> Theme {
    Theme::default()
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
