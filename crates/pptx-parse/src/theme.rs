use ooxml_drawingml::{
    ColorMap, Theme, ThemeColorScheme, ThemeFont, ThemeFontScheme, ThemeFormatScheme,
};

use crate::drawing::{
    parse_effects, parse_fill_element, parse_outline_element, picture_fill_element,
};
use crate::model::PictureFill;
use crate::relationships::Relationship;
use crate::xml::XmlElement;

pub(crate) fn parse_theme(root: &XmlElement) -> Theme {
    let elements = root.child("themeElements");
    Theme {
        name: root.attribute("name").unwrap_or("Office Theme").to_owned(),
        color_scheme: parse_color_scheme(elements.and_then(|value| value.child("clrScheme"))),
        font_scheme: parse_font_scheme(elements.and_then(|value| value.child("fontScheme"))),
        color_map: ColorMap::default(),
    }
}

pub(crate) fn parse_format_scheme(root: &XmlElement) -> ThemeFormatScheme {
    let Some(element) = root
        .child("themeElements")
        .and_then(|value| value.child("fmtScheme"))
    else {
        return ThemeFormatScheme::default();
    };
    ThemeFormatScheme {
        background_fills: style_list(element.child("bgFillStyleLst"), parse_fill_element),
        fills: style_list(element.child("fillStyleLst"), parse_fill_element),
        lines: style_list(element.child("lnStyleLst"), parse_outline_element),
        effects: style_list(element.child("effectStyleLst"), parse_effects),
    }
}

/// `a:bgFillStyleLst` picture entries, aligned with [`parse_format_scheme`].
pub(crate) fn parse_background_pictures(
    root: &XmlElement,
    relationships: &[Relationship],
) -> Vec<Option<PictureFill>> {
    let Some(element) = root
        .child("themeElements")
        .and_then(|value| value.child("fmtScheme"))
        .and_then(|value| value.child("bgFillStyleLst"))
    else {
        return Vec::new();
    };
    let pictures: Vec<Option<PictureFill>> = element
        .child_elements()
        .map(|fill| picture_fill_element(fill, relationships))
        .collect();
    if pictures.iter().all(Option::is_none) {
        Vec::new()
    } else {
        pictures
    }
}

/// Preserves unsupported entries as empty slots.
fn style_list<T>(
    element: Option<&XmlElement>,
    parse: impl Fn(&XmlElement) -> Option<T>,
) -> Vec<Option<T>> {
    element
        .into_iter()
        .flat_map(XmlElement::child_elements)
        .map(parse)
        .collect()
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
    let rgb = match element.local_name() {
        "srgbClr" => element.attribute("val").map(str::to_owned),
        "sysClr" => system_color_value(element).map(str::to_owned),
        _ => None,
    }?;
    let transforms: Vec<_> = element
        .child_elements()
        .filter_map(|child| {
            ooxml_drawingml::ColorTransform::from_xml(child.local_name(), child.attribute("val"))
        })
        .collect();
    if transforms.is_empty() {
        Some(rgb)
    } else {
        ooxml_drawingml::resolve_color_value_to_hex(Some(&ooxml_drawingml::ColorValue {
            rgb: Some(rgb),
            transforms,
            ..Default::default()
        }))
        .map(|hex| hex.trim_start_matches('#').to_owned())
    }
}

pub(crate) fn system_color_value(element: &XmlElement) -> Option<&str> {
    let value = element.attribute("val");
    match value {
        Some("window" | "windowText") => system_color(value),
        _ => element.attribute("lastClr").or_else(|| system_color(value)),
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
    fn theme_effect_styles_preserve_indices_and_shadow_colors() {
        let limits = ParseLimits::default();
        let root = parse_xml(
            br#"<a:theme><a:themeElements><a:fmtScheme><a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="76200" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="38100" dist="76200" dir="0"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:outerShdw></a:effectLst></a:effectStyle>
</a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>"#,
            "ppt/theme/theme1.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let scheme = parse_format_scheme(&root);
        assert!(!scheme.is_empty());
        assert_eq!(scheme.effects.len(), 3);
        assert_eq!(
            scheme.effects[0],
            Some(ooxml_drawingml::ShapeEffects::default())
        );
        for (index, blur, distance, direction, rgb, theme_color, alpha) in [
            (1, 76200, 38100, 5400000, Some("000000"), None, 0.38),
            (2, 38100, 76200, 0, None, Some("phClr"), 0.5),
        ] {
            let shadow = scheme.effects[index]
                .as_ref()
                .unwrap()
                .outer_shadow
                .as_ref()
                .unwrap();
            assert_eq!(shadow.blur_radius, blur);
            assert_eq!(shadow.distance, distance);
            assert_eq!(shadow.direction, direction);
            let color = shadow.color.as_ref().unwrap();
            assert_eq!(color.rgb.as_deref(), rgb);
            assert_eq!(color.theme_color.as_deref(), theme_color);
            assert_eq!(ooxml_drawingml::resolve_color_value_alpha(color), alpha);
        }
    }

    #[test]
    fn theme_effect_styles_preserve_unsupported_slots() {
        let limits = ParseLimits::default();
        let root = parse_xml(
            br#"<a:theme><a:themeElements><a:fmtScheme><a:effectStyleLst><a:effectStyle><a:effectDag/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>"#,
            "ppt/theme/theme1.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(
            parse_format_scheme(&root).effects,
            vec![None, Some(ooxml_drawingml::ShapeEffects::default())]
        );
    }

    #[test]
    fn theme_color_definitions_apply_ordered_transforms() {
        let limits = ParseLimits::default();
        let root = parse_xml(
            br#"<a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="4472C4"><a:tint val="50000"/><a:lumMod val="50000"/></a:srgbClr></a:accent1><a:accent2><a:sysClr lastClr="4472C4"><a:lumMod val="50000"/><a:tint val="50000"/></a:sysClr></a:accent2></a:clrScheme></a:themeElements></a:theme>"#,
            "ppt/theme/theme1.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let theme = parse_theme(&root);
        assert_eq!(theme.color_scheme.accent1, "3F5493");
        assert_eq!(theme.color_scheme.accent2, "BDBFC6");
    }

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

    #[test]
    fn window_system_colors_ignore_the_saving_machine_palette() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<a:theme name="Dark"><a:themeElements><a:clrScheme><a:dk1><a:sysClr val="windowText" lastClr="FFFFFF"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="202020"/></a:lt1><a:accent1><a:sysClr val="highlight" lastClr="3399FF"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#,
            "ppt/theme/theme1.xml",
            &mut budget,
        )
        .unwrap();
        let theme = parse_theme(&root);
        assert_eq!(theme.color_scheme.dk1, "000000");
        assert_eq!(theme.color_scheme.lt1, "FFFFFF");
        assert_eq!(theme.color_scheme.accent1, "3399FF");
    }
}
