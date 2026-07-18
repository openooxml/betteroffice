//! Typed run, paragraph, row, cell, and table property bags shared by styles and numbering.

use serde::{Deserialize, Serialize};

use crate::borders::{BorderSpec, Borders, parse_border_spec, parse_paragraph_borders};
use crate::scalars::{
    ColorValue, ShadingProperties, UnderlineValue, parse_color_value, parse_shading_properties,
    parse_underline,
};
use crate::tabs::TabStop;
use crate::theme::{Theme, resolve_theme_font_ref};
use crate::xml::XmlElement;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontFamily {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ascii: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h_ansi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ascii_theme: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h_ansi_theme: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asia_theme: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs_theme: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hint: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunLanguage {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub latin: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextGlow {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub radius: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextShadow {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blur_radius: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub direction: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextReflection {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blur_radius: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end_opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub direction: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct TextGradientStop {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextFill {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub angle: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stops: Option<Vec<TextGradientStop>>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextOutline {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_fill: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextModernEffects {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub glow: Option<TextGlow>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shadow: Option<TextShadow>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reflection: Option<TextReflection>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_fill: Option<TextFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_outline: Option<TextOutline>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub double_strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
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
    pub font_family: Option<FontFamily>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub language: Option<RunLanguage>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kerning: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effect: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emboss: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub imprint: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shadow: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modern_effects: Option<TextModernEffects>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

pub fn parse_run_properties(
    r_pr: Option<&XmlElement>,
    theme: Option<&Theme>,
) -> Option<TextFormatting> {
    let r_pr = r_pr?;
    let mut value = TextFormatting::default();
    value.bold = boolean_child(r_pr, "b");
    value.bold_cs = boolean_child(r_pr, "bCs");
    value.italic = boolean_child(r_pr, "i");
    value.italic_cs = boolean_child(r_pr, "iCs");
    value.underline = parse_underline(r_pr.child("w", "u"));
    value.strike = boolean_child(r_pr, "strike");
    value.double_strike = boolean_child(r_pr, "dstrike");
    value.vert_align =
        string_child_in(r_pr, "vertAlign", &["superscript", "subscript", "baseline"]);
    value.small_caps = boolean_child(r_pr, "smallCaps");
    value.all_caps = boolean_child(r_pr, "caps");
    value.hidden = boolean_child(r_pr, "vanish");
    value.color = r_pr.child("w", "color").map(|element| {
        parse_color_value(
            element.attribute(Some("w"), "val"),
            element.attribute(Some("w"), "themeColor"),
            element.attribute(Some("w"), "themeTint"),
            element.attribute(Some("w"), "themeShade"),
        )
    });
    value.highlight = string_child(r_pr, "highlight", "val", true);
    let shading_element = r_pr.child("w", "shd");
    value.shading = parse_shading_properties(shading_element);
    value.font_size = numeric_child(r_pr, "sz", "val");
    value.font_size_cs = numeric_child(r_pr, "szCs", "val");
    value.font_family = r_pr.child("w", "rFonts").map(|fonts| {
        let mut family = FontFamily {
            ascii: attribute_owned(fonts, "ascii"),
            h_ansi: attribute_owned(fonts, "hAnsi"),
            east_asia: attribute_owned(fonts, "eastAsia"),
            cs: attribute_owned(fonts, "cs"),
            ..FontFamily::default()
        };
        family.hint = fonts
            .attribute(Some("w"), "hint")
            .filter(|hint| matches!(*hint, "default" | "eastAsia" | "cs"))
            .map(str::to_owned);
        family.ascii_theme = attribute_nonempty(fonts, "asciiTheme");
        family.h_ansi_theme = attribute_nonempty(fonts, "hAnsiTheme");
        family.east_asia_theme = attribute_nonempty(fonts, "eastAsiaTheme");
        // Incumbent spelling is deliberately lower-case `cstheme`.
        family.cs_theme = attribute_nonempty(fonts, "cstheme");
        if let Some(reference) = family.ascii_theme.as_deref()
            && theme.is_some()
            && family.ascii.as_deref().is_none_or(str::is_empty)
        {
            family.ascii = Some(resolve_theme_font_ref(theme, reference));
        }
        if let Some(reference) = family.h_ansi_theme.as_deref()
            && theme.is_some()
            && family.h_ansi.as_deref().is_none_or(str::is_empty)
        {
            family.h_ansi = Some(resolve_theme_font_ref(theme, reference));
        }
        if let Some(reference) = family.east_asia_theme.as_deref()
            && theme.is_some()
            && family.east_asia.as_deref().is_none_or(str::is_empty)
        {
            family.east_asia = Some(resolve_theme_font_ref(theme, reference));
        }
        if let Some(reference) = family.cs_theme.as_deref()
            && theme.is_some()
            && family.cs.as_deref().is_none_or(str::is_empty)
        {
            family.cs = Some(resolve_theme_font_ref(theme, reference));
        }
        family
    });
    value.spacing = numeric_child(r_pr, "spacing", "val");
    value.position = numeric_child(r_pr, "position", "val");
    value.scale = numeric_child(r_pr, "w", "val");
    value.kerning = numeric_child(r_pr, "kern", "val");
    value.effect = string_child(r_pr, "effect", "val", true);
    value.emphasis_mark = string_child(r_pr, "em", "val", true);
    value.emboss = boolean_child(r_pr, "emboss");
    value.imprint = boolean_child(r_pr, "imprint");
    value.outline = boolean_child(r_pr, "outline");
    value.shadow = boolean_child(r_pr, "shadow");
    value.rtl = boolean_child(r_pr, "rtl");
    value.cs = boolean_child(r_pr, "cs");
    value.language = r_pr.child("w", "lang").and_then(|language| {
        let value = RunLanguage {
            latin: valid_language_tag(language.attribute(Some("w"), "val")),
            east_asia: valid_language_tag(language.attribute(Some("w"), "eastAsia")),
            bidi: valid_language_tag(language.attribute(Some("w"), "bidi")),
        };
        (value != RunLanguage::default()).then_some(value)
    });
    value.style_id = string_child(r_pr, "rStyle", "val", true);
    (value != TextFormatting::default() || shading_element.is_some()).then_some(value)
}

fn valid_language_tag(raw: Option<&str>) -> Option<String> {
    let value = raw?.trim();
    if value.is_empty()
        || value.chars().count() > 63
        || value.split('-').any(|part| {
            part.is_empty()
                || part.len() > 8
                || !part.as_bytes().iter().all(u8::is_ascii_alphanumeric)
        })
    {
        return None;
    }
    Some(value.to_owned())
}

pub fn merge_text_formatting(
    target: Option<&TextFormatting>,
    source: Option<&TextFormatting>,
) -> Option<TextFormatting> {
    match (target, source) {
        (None, None) => None,
        (Some(target), None) => Some(target.clone()),
        (None, Some(source)) => Some(source.clone()),
        (Some(target), Some(source)) => {
            let mut result = target.clone();
            overlay(&mut result.bold, &source.bold);
            overlay(&mut result.bold_cs, &source.bold_cs);
            overlay(&mut result.italic, &source.italic);
            overlay(&mut result.italic_cs, &source.italic_cs);
            if let Some(source) = &source.underline {
                result.underline = Some(UnderlineValue {
                    style: source.style.clone(),
                    color: source.color.clone().or_else(|| {
                        target
                            .underline
                            .as_ref()
                            .and_then(|value| value.color.clone())
                    }),
                });
            }
            overlay(&mut result.strike, &source.strike);
            overlay(&mut result.double_strike, &source.double_strike);
            overlay(&mut result.vert_align, &source.vert_align);
            overlay(&mut result.small_caps, &source.small_caps);
            overlay(&mut result.all_caps, &source.all_caps);
            overlay(&mut result.hidden, &source.hidden);
            if let Some(color) = &source.color {
                let explicit = color.rgb.is_some()
                    || color.theme_color.is_some()
                    || color.theme_tint.is_some()
                    || color.theme_shade.is_some();
                if color.auto != Some(true) || explicit {
                    result.color = Some(color.clone());
                }
            }
            overlay(&mut result.highlight, &source.highlight);
            if let Some(source) = &source.shading {
                let mut merged = target.shading.clone().unwrap_or_default();
                overlay(&mut merged.color, &source.color);
                overlay(&mut merged.fill, &source.fill);
                overlay(&mut merged.pattern, &source.pattern);
                result.shading = Some(merged);
            }
            overlay(&mut result.font_size, &source.font_size);
            overlay(&mut result.font_size_cs, &source.font_size_cs);
            if let Some(source) = &source.font_family {
                result.font_family = Some(merge_font_family(target.font_family.as_ref(), source));
            }
            if let Some(source) = &source.language {
                let mut merged = target.language.clone().unwrap_or_default();
                overlay(&mut merged.latin, &source.latin);
                overlay(&mut merged.east_asia, &source.east_asia);
                overlay(&mut merged.bidi, &source.bidi);
                result.language = Some(merged);
            }
            overlay(&mut result.spacing, &source.spacing);
            overlay(&mut result.position, &source.position);
            overlay(&mut result.scale, &source.scale);
            overlay(&mut result.kerning, &source.kerning);
            overlay(&mut result.effect, &source.effect);
            overlay(&mut result.emphasis_mark, &source.emphasis_mark);
            overlay(&mut result.emboss, &source.emboss);
            overlay(&mut result.imprint, &source.imprint);
            overlay(&mut result.outline, &source.outline);
            overlay(&mut result.shadow, &source.shadow);
            overlay(&mut result.modern_effects, &source.modern_effects);
            overlay(&mut result.rtl, &source.rtl);
            overlay(&mut result.cs, &source.cs);
            overlay(&mut result.style_id, &source.style_id);
            Some(result)
        }
    }
}

fn merge_font_family(target: Option<&FontFamily>, source: &FontFamily) -> FontFamily {
    let mut result = target.cloned().unwrap_or_default();
    replace_font_pair(
        &mut result.ascii,
        &mut result.ascii_theme,
        &source.ascii,
        &source.ascii_theme,
    );
    replace_font_pair(
        &mut result.h_ansi,
        &mut result.h_ansi_theme,
        &source.h_ansi,
        &source.h_ansi_theme,
    );
    replace_font_pair(
        &mut result.east_asia,
        &mut result.east_asia_theme,
        &source.east_asia,
        &source.east_asia_theme,
    );
    replace_font_pair(
        &mut result.cs,
        &mut result.cs_theme,
        &source.cs,
        &source.cs_theme,
    );
    overlay(&mut result.hint, &source.hint);
    result
}

fn replace_font_pair(
    explicit: &mut Option<String>,
    theme: &mut Option<String>,
    source_explicit: &Option<String>,
    source_theme: &Option<String>,
) {
    if source_explicit.is_some() || source_theme.is_some() {
        explicit.clone_from(source_explicit);
        theme.clone_from(source_theme);
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberingProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ilvl: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct SpacingExplicit {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphFrame {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_before: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_after: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing_rule: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before_autospacing: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after_autospacing: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_explicit: Option<SpacingExplicit>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent_right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent_first_line: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hanging_indent: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<Borders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<ShadingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Vec<TabStop>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub contextual_spacing: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<NumberingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_pr_from_style: Option<NumberingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frame: Option<ParagraphFrame>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suppress_line_numbers: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suppress_auto_hyphens: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_properties: Option<TextFormatting>,
}

pub fn parse_paragraph_properties(
    p_pr: Option<&XmlElement>,
    theme: Option<&Theme>,
) -> Option<ParagraphFormatting> {
    let p_pr = p_pr?;
    let mut value = ParagraphFormatting::default();
    value.alignment = string_child(p_pr, "jc", "val", true);
    value.bidi = boolean_child(p_pr, "bidi");
    if let Some(spacing) = p_pr.child("w", "spacing") {
        value.space_before = spacing.parse_numeric_attribute(Some("w"), "before", 1.0);
        value.space_after = spacing.parse_numeric_attribute(Some("w"), "after", 1.0);
        value.line_spacing = spacing.parse_numeric_attribute(Some("w"), "line", 1.0);
        value.line_spacing_rule = attribute_nonempty(spacing, "lineRule");
        value.before_autospacing = spacing
            .attribute(Some("w"), "beforeAutospacing")
            .filter(|raw| !raw.is_empty())
            .map(|raw| matches!(raw, "1" | "true"));
        value.after_autospacing = spacing
            .attribute(Some("w"), "afterAutospacing")
            .filter(|raw| !raw.is_empty())
            .map(|raw| matches!(raw, "1" | "true"));
    }
    if let Some(indent) = p_pr.child("w", "ind") {
        value.indent_left = indent.parse_numeric_attribute(Some("w"), "left", 1.0);
        value.indent_right = indent.parse_numeric_attribute(Some("w"), "right", 1.0);
        value.indent_first_line = indent.parse_numeric_attribute(Some("w"), "firstLine", 1.0);
        if let Some(hanging) = indent.parse_numeric_attribute(Some("w"), "hanging", 1.0) {
            value.indent_first_line = Some(-hanging);
            value.hanging_indent = Some(true);
        }
    }
    value.borders = parse_paragraph_borders(p_pr.child("w", "pBdr"));
    let shading_element = p_pr.child("w", "shd");
    value.shading = parse_shading_properties(shading_element);
    let tabs_element = p_pr.child("w", "tabs");
    value.tabs = parse_style_tab_stops(tabs_element);
    value.keep_next = boolean_child(p_pr, "keepNext");
    value.keep_lines = boolean_child(p_pr, "keepLines");
    value.widow_control = boolean_child(p_pr, "widowControl");
    value.page_break_before = boolean_child(p_pr, "pageBreakBefore");
    value.contextual_spacing = boolean_child(p_pr, "contextualSpacing");
    if let Some(num_pr) = p_pr.child("w", "numPr") {
        let num_id_element = num_pr.child("w", "numId");
        let ilvl_element = num_pr.child("w", "ilvl");
        if num_id_element.is_some() || ilvl_element.is_some() {
            value.num_pr = Some(NumberingProperties {
                num_id: num_id_element
                    .and_then(|element| element.parse_numeric_attribute(Some("w"), "val", 1.0)),
                ilvl: ilvl_element
                    .and_then(|element| element.parse_numeric_attribute(Some("w"), "val", 1.0)),
            });
        }
    }
    value.outline_level = numeric_child(p_pr, "outlineLvl", "val");
    value.style_id = string_child(p_pr, "pStyle", "val", true);
    value.suppress_line_numbers = boolean_child(p_pr, "suppressLineNumbers");
    value.suppress_auto_hyphens = boolean_child(p_pr, "suppressAutoHyphens");
    let run_properties_element = p_pr.child("w", "rPr");
    value.run_properties = parse_run_properties(run_properties_element, theme);
    (value != ParagraphFormatting::default()
        || shading_element.is_some()
        || tabs_element.is_some()
        || run_properties_element.is_some())
    .then_some(value)
}

fn parse_style_tab_stops(tabs: Option<&XmlElement>) -> Option<Vec<TabStop>> {
    let tabs = tabs?;
    let stops: Vec<_> = tabs
        .children_named("w", "tab")
        .filter_map(|tab| {
            Some(TabStop {
                position: tab.parse_numeric_attribute(Some("w"), "pos", 1.0)?,
                alignment: tab.attribute(Some("w"), "val")?.to_owned(),
                leader: attribute_nonempty(tab, "leader"),
            })
        })
        .collect();
    (!stops.is_empty()).then_some(stops)
}

pub fn merge_paragraph_formatting(
    target: Option<&ParagraphFormatting>,
    source: Option<&ParagraphFormatting>,
) -> Option<ParagraphFormatting> {
    match (target, source) {
        (None, None) => None,
        (Some(target), None) => Some(target.clone()),
        (None, Some(source)) => Some(source.clone()),
        (Some(target), Some(source)) => {
            let mut result = target.clone();
            overlay(&mut result.alignment, &source.alignment);
            overlay(&mut result.bidi, &source.bidi);
            overlay(&mut result.space_before, &source.space_before);
            overlay(&mut result.space_after, &source.space_after);
            overlay(&mut result.line_spacing, &source.line_spacing);
            overlay(&mut result.line_spacing_rule, &source.line_spacing_rule);
            overlay(&mut result.before_autospacing, &source.before_autospacing);
            overlay(&mut result.after_autospacing, &source.after_autospacing);
            overlay(&mut result.spacing_explicit, &source.spacing_explicit);
            overlay(&mut result.indent_left, &source.indent_left);
            overlay(&mut result.indent_right, &source.indent_right);
            overlay(&mut result.indent_first_line, &source.indent_first_line);
            overlay(&mut result.hanging_indent, &source.hanging_indent);
            if let Some(source) = &source.borders {
                result.borders = Some(merge_borders(target.borders.as_ref(), source));
            }
            overlay(&mut result.shading, &source.shading);
            overlay(&mut result.tabs, &source.tabs);
            overlay(&mut result.keep_next, &source.keep_next);
            overlay(&mut result.keep_lines, &source.keep_lines);
            overlay(&mut result.widow_control, &source.widow_control);
            overlay(&mut result.page_break_before, &source.page_break_before);
            overlay(&mut result.contextual_spacing, &source.contextual_spacing);
            if let Some(source) = &source.num_pr {
                let mut merged = target.num_pr.clone().unwrap_or_default();
                overlay(&mut merged.num_id, &source.num_id);
                overlay(&mut merged.ilvl, &source.ilvl);
                result.num_pr = Some(merged);
            }
            overlay(&mut result.num_pr_from_style, &source.num_pr_from_style);
            overlay(&mut result.outline_level, &source.outline_level);
            overlay(&mut result.style_id, &source.style_id);
            overlay(&mut result.frame, &source.frame);
            overlay(
                &mut result.suppress_line_numbers,
                &source.suppress_line_numbers,
            );
            overlay(
                &mut result.suppress_auto_hyphens,
                &source.suppress_auto_hyphens,
            );
            result.run_properties = merge_text_formatting(
                target.run_properties.as_ref(),
                source.run_properties.as_ref(),
            );
            Some(result)
        }
    }
}

fn merge_borders(target: Option<&Borders>, source: &Borders) -> Borders {
    let mut result = target.cloned().unwrap_or_default();
    overlay(&mut result.top, &source.top);
    overlay(&mut result.bottom, &source.bottom);
    overlay(&mut result.left, &source.left);
    overlay(&mut result.right, &source.right);
    overlay(&mut result.inside_h, &source.inside_h);
    overlay(&mut result.inside_v, &source.inside_v);
    overlay(&mut result.between, &source.between);
    overlay(&mut result.bar, &source.bar);
    overlay(&mut result.start, &source.start);
    overlay(&mut result.end, &source.end);
    overlay(&mut result.tl2br, &source.tl2br);
    overlay(&mut result.tr2bl, &source.tr2bl);
    result
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TableMeasurement {
    pub value: f64,
    #[serde(rename = "type")]
    pub kind: String,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct CellMargins {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end: Option<TableMeasurement>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableLook {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_row: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_column: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_row: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_h_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_v_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_row: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_row: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_column: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub odd_h_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_h_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub odd_v_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_v_band: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub nw_cell: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ne_cell: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sw_cell: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub se_cell: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FloatingTableProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horz_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_x_spec: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_y_spec: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_from_text: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub justification: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_spacing: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<Borders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_margins: Option<CellMargins>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub look: Option<TableLook>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<ShadingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub overlap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub floating: Option<FloatingTableProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_row_band_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_col_band_size: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRowFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cant_split: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub justification: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub conditional_format: Option<ConditionalFormatStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_before: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_after: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_before: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_after: Option<TableMeasurement>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<TableMeasurement>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<Borders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<CellMargins>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<ShadingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_merge: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hide_mark: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub conditional_format: Option<ConditionalFormatStyle>,
}

pub fn parse_table_properties(element: Option<&XmlElement>) -> Option<TableFormatting> {
    let element = element?;
    let mut value = TableFormatting::default();
    let width_element = element.child("w", "tblW");
    value.width = parse_table_measurement(width_element);
    value.justification = string_child_in(element, "jc", &["left", "center", "right"]);
    let cell_spacing_element = element.child("w", "tblCellSpacing");
    value.cell_spacing = parse_table_measurement(cell_spacing_element);
    let indent_element = element.child("w", "tblInd");
    value.indent = parse_table_measurement(indent_element);
    let borders_element = element.child("w", "tblBorders");
    value.borders = parse_style_table_borders(borders_element);
    let cell_margins_element = element.child("w", "tblCellMar");
    value.cell_margins = parse_cell_margins(cell_margins_element);
    value.layout = element
        .child("w", "tblLayout")
        .and_then(|child| child.attribute(Some("w"), "type"))
        .filter(|raw| matches!(*raw, "fixed" | "autofit"))
        .map(str::to_owned);
    value.style_id = string_child(element, "tblStyle", "val", true);
    value.style_row_band_size = numeric_child(element, "tblStyleRowBandSize", "val")
        .filter(|size| *size > 0.0 && *size <= 1024.0);
    value.style_col_band_size = numeric_child(element, "tblStyleColBandSize", "val")
        .filter(|size| *size > 0.0 && *size <= 1024.0);
    let look_element = element.child("w", "tblLook");
    value.look = parse_table_look(look_element);
    let shading_element = element.child("w", "shd");
    value.shading = parse_shading_properties(shading_element);
    value.bidi = boolean_child(element, "bidiVisual");
    (value != TableFormatting::default()
        || width_element.is_some()
        || cell_spacing_element.is_some()
        || indent_element.is_some()
        || borders_element.is_some()
        || cell_margins_element.is_some()
        || look_element.is_some()
        || shading_element.is_some())
    .then_some(value)
}

pub fn parse_table_row_properties(element: Option<&XmlElement>) -> Option<TableRowFormatting> {
    let element = element?;
    let mut value = TableRowFormatting::default();
    let height_element = element.child("w", "trHeight");
    if let Some(height) = height_element {
        value.height = parse_table_measurement(Some(height));
        value.height_rule = attribute_nonempty(height, "hRule");
    }
    value.header = boolean_child(element, "tblHeader");
    value.cant_split = boolean_child(element, "cantSplit");
    value.justification = string_child_in(element, "jc", &["left", "center", "right"]);
    value.hidden = boolean_child(element, "hidden");
    value.grid_before = numeric_child(element, "gridBefore", "val")
        .filter(|count| *count >= 0.0 && *count <= 16_384.0);
    value.grid_after = numeric_child(element, "gridAfter", "val")
        .filter(|count| *count >= 0.0 && *count <= 16_384.0);
    value.width_before = parse_table_measurement(element.child("w", "wBefore"));
    value.width_after = parse_table_measurement(element.child("w", "wAfter"));
    (value != TableRowFormatting::default() || height_element.is_some()).then_some(value)
}

pub fn parse_table_cell_properties(element: Option<&XmlElement>) -> Option<TableCellFormatting> {
    let element = element?;
    let mut value = TableCellFormatting::default();
    let width_element = element.child("w", "tcW");
    value.width = parse_table_measurement(width_element);
    let borders_element = element.child("w", "tcBorders");
    value.borders = parse_style_table_borders(borders_element);
    let margins_element = element.child("w", "tcMar");
    value.margins = parse_cell_margins(margins_element);
    let shading_element = element.child("w", "shd");
    value.shading = parse_shading_properties(shading_element);
    value.vertical_align = string_child_in(element, "vAlign", &["top", "center", "bottom"]);
    value.text_direction = string_child(element, "textDirection", "val", true);
    value.grid_span = numeric_child(element, "gridSpan", "val");
    value.v_merge = element.child("w", "vMerge").map(|merge| {
        if merge.attribute(Some("w"), "val") == Some("restart") {
            "restart"
        } else {
            "continue"
        }
        .to_owned()
    });
    value.fit_text = boolean_child(element, "tcFitText");
    value.no_wrap = boolean_child(element, "noWrap");
    value.hide_mark = boolean_child(element, "hideMark");
    (value != TableCellFormatting::default()
        || width_element.is_some()
        || borders_element.is_some()
        || margins_element.is_some()
        || shading_element.is_some())
    .then_some(value)
}

fn parse_table_measurement(element: Option<&XmlElement>) -> Option<TableMeasurement> {
    let element = element?;
    Some(TableMeasurement {
        value: element.parse_numeric_attribute(Some("w"), "w", 1.0)?,
        kind: element
            .attribute(Some("w"), "type")
            .filter(|value| !value.is_empty())?
            .to_owned(),
    })
}

fn parse_cell_margins(element: Option<&XmlElement>) -> Option<CellMargins> {
    let element = element?;
    let value = CellMargins {
        top: parse_table_measurement(element.child("w", "top")),
        bottom: parse_table_measurement(element.child("w", "bottom")),
        left: parse_table_measurement(element.child("w", "left")),
        right: parse_table_measurement(element.child("w", "right")),
        start: parse_table_measurement(element.child("w", "start")),
        end: parse_table_measurement(element.child("w", "end")),
    };
    (value != CellMargins::default()).then_some(value)
}

fn parse_style_table_borders(element: Option<&XmlElement>) -> Option<Borders> {
    let element = element?;
    let value = Borders {
        top: parse_border_spec(element.child("w", "top")),
        bottom: parse_border_spec(element.child("w", "bottom")),
        left: parse_border_spec(element.child("w", "left")),
        right: parse_border_spec(element.child("w", "right")),
        inside_h: parse_border_spec(element.child("w", "insideH")),
        inside_v: parse_border_spec(element.child("w", "insideV")),
        start: parse_border_spec(element.child("w", "start")),
        end: parse_border_spec(element.child("w", "end")),
        tl2br: parse_border_spec(element.child("w", "tl2br")),
        tr2bl: parse_border_spec(element.child("w", "tr2bl")),
        ..Borders::default()
    };
    (value != Borders::default()).then_some(value)
}

fn parse_table_look(element: Option<&XmlElement>) -> Option<TableLook> {
    let element = element?;
    let mut value = TableLook::default();
    if let Some(raw) = element
        .attribute(Some("w"), "val")
        .filter(|raw| !raw.is_empty())
    {
        value.value = Some(raw.chars().take(8).collect());
        if let Some(bits) = parse_hex_prefix(raw) {
            value.first_row = Some(bits & 0x0020 != 0);
            value.last_row = Some(bits & 0x0040 != 0);
            value.first_column = Some(bits & 0x0080 != 0);
            value.last_column = Some(bits & 0x0100 != 0);
            value.no_h_band = Some(bits & 0x0200 != 0);
            value.no_v_band = Some(bits & 0x0400 != 0);
        }
    }
    for (attribute, slot) in [
        ("firstColumn", &mut value.first_column),
        ("firstRow", &mut value.first_row),
        ("lastColumn", &mut value.last_column),
        ("lastRow", &mut value.last_row),
        ("noHBand", &mut value.no_h_band),
        ("noVBand", &mut value.no_v_band),
    ] {
        if let Some(raw) = element.attribute(Some("w"), attribute) {
            *slot = Some(!matches_ci(raw, &["0", "false", "off"]));
        }
    }
    (value != TableLook::default()).then_some(value)
}

fn parse_hex_prefix(raw: &str) -> Option<u32> {
    let trimmed = raw.trim_start();
    let unsigned = trimmed
        .strip_prefix('+')
        .or_else(|| trimmed.strip_prefix('-'))
        .unwrap_or(trimmed);
    let unsigned = unsigned
        .strip_prefix("0x")
        .or_else(|| unsigned.strip_prefix("0X"))
        .unwrap_or(unsigned);
    let digits: String = unsigned
        .chars()
        .take_while(|character| character.is_ascii_hexdigit())
        .collect();
    if digits.is_empty() {
        None
    } else {
        u64::from_str_radix(&digits.chars().take(16).collect::<String>(), 16)
            .ok()
            .map(|value| value as u32)
    }
}

pub fn merge_table_formatting(
    target: Option<&TableFormatting>,
    source: Option<&TableFormatting>,
) -> Option<TableFormatting> {
    match (target, source) {
        (None, None) => None,
        (Some(target), None) => Some(target.clone()),
        (None, Some(source)) => Some(source.clone()),
        (Some(target), Some(source)) => {
            let mut value = target.clone();
            overlay(&mut value.width, &source.width);
            overlay(&mut value.justification, &source.justification);
            overlay(&mut value.cell_spacing, &source.cell_spacing);
            overlay(&mut value.indent, &source.indent);
            if let Some(source) = &source.borders {
                value.borders = Some(merge_table_borders(target.borders.as_ref(), source));
            }
            if let Some(source) = &source.cell_margins {
                let mut merged = target.cell_margins.clone().unwrap_or_default();
                overlay(&mut merged.top, &source.top);
                overlay(&mut merged.bottom, &source.bottom);
                overlay(&mut merged.left, &source.left);
                overlay(&mut merged.right, &source.right);
                overlay(&mut merged.start, &source.start);
                overlay(&mut merged.end, &source.end);
                value.cell_margins = Some(merged);
            }
            overlay(&mut value.layout, &source.layout);
            overlay(&mut value.style_id, &source.style_id);
            if let Some(source) = &source.look {
                let mut merged = target.look.clone().unwrap_or_default();
                overlay(&mut merged.first_column, &source.first_column);
                overlay(&mut merged.first_row, &source.first_row);
                overlay(&mut merged.last_column, &source.last_column);
                overlay(&mut merged.last_row, &source.last_row);
                overlay(&mut merged.no_h_band, &source.no_h_band);
                overlay(&mut merged.no_v_band, &source.no_v_band);
                overlay(&mut merged.value, &source.value);
                value.look = Some(merged);
            }
            if let Some(source) = &source.shading {
                value.shading = Some(merge_shading_deep(target.shading.as_ref(), source));
            }
            overlay(&mut value.overlap, &source.overlap);
            overlay(&mut value.floating, &source.floating);
            overlay(&mut value.bidi, &source.bidi);
            overlay(&mut value.style_row_band_size, &source.style_row_band_size);
            overlay(&mut value.style_col_band_size, &source.style_col_band_size);
            Some(value)
        }
    }
}

pub fn merge_row_formatting(
    target: Option<&TableRowFormatting>,
    source: Option<&TableRowFormatting>,
) -> Option<TableRowFormatting> {
    match (target, source) {
        (None, None) => None,
        (Some(target), None) => Some(target.clone()),
        (None, Some(source)) => Some(source.clone()),
        (Some(target), Some(source)) => {
            let mut value = target.clone();
            overlay(&mut value.height, &source.height);
            overlay(&mut value.height_rule, &source.height_rule);
            overlay(&mut value.header, &source.header);
            overlay(&mut value.cant_split, &source.cant_split);
            overlay(&mut value.justification, &source.justification);
            overlay(&mut value.hidden, &source.hidden);
            overlay(&mut value.conditional_format, &source.conditional_format);
            overlay(&mut value.grid_before, &source.grid_before);
            overlay(&mut value.grid_after, &source.grid_after);
            overlay(&mut value.width_before, &source.width_before);
            overlay(&mut value.width_after, &source.width_after);
            Some(value)
        }
    }
}

pub fn merge_cell_formatting(
    target: Option<&TableCellFormatting>,
    source: Option<&TableCellFormatting>,
) -> Option<TableCellFormatting> {
    match (target, source) {
        (None, None) => None,
        (Some(target), None) => Some(target.clone()),
        (None, Some(source)) => Some(source.clone()),
        (Some(target), Some(source)) => {
            let mut value = target.clone();
            overlay(&mut value.width, &source.width);
            if let Some(source) = &source.borders {
                value.borders = Some(merge_table_borders(target.borders.as_ref(), source));
            }
            if let Some(source) = &source.margins {
                let mut merged = target.margins.clone().unwrap_or_default();
                overlay(&mut merged.top, &source.top);
                overlay(&mut merged.bottom, &source.bottom);
                overlay(&mut merged.left, &source.left);
                overlay(&mut merged.right, &source.right);
                overlay(&mut merged.start, &source.start);
                overlay(&mut merged.end, &source.end);
                value.margins = Some(merged);
            }
            if let Some(source) = &source.shading {
                value.shading = Some(merge_shading_deep(target.shading.as_ref(), source));
            }
            overlay(&mut value.vertical_align, &source.vertical_align);
            overlay(&mut value.text_direction, &source.text_direction);
            overlay(&mut value.grid_span, &source.grid_span);
            overlay(&mut value.v_merge, &source.v_merge);
            overlay(&mut value.fit_text, &source.fit_text);
            overlay(&mut value.no_wrap, &source.no_wrap);
            overlay(&mut value.hide_mark, &source.hide_mark);
            overlay(&mut value.conditional_format, &source.conditional_format);
            Some(value)
        }
    }
}

fn boolean_child(parent: &XmlElement, name: &str) -> Option<bool> {
    parent
        .child("w", name)
        .map(|element| element.parse_boolean("w"))
}

fn numeric_child(parent: &XmlElement, child: &str, attribute: &str) -> Option<f64> {
    parent
        .child("w", child)
        .and_then(|element| element.parse_numeric_attribute(Some("w"), attribute, 1.0))
}

fn string_child(
    parent: &XmlElement,
    child: &str,
    attribute: &str,
    nonempty: bool,
) -> Option<String> {
    let value = parent.child("w", child)?.attribute(Some("w"), attribute)?;
    (!nonempty || !value.is_empty()).then(|| value.to_owned())
}

fn string_child_in(parent: &XmlElement, child: &str, allowed: &[&str]) -> Option<String> {
    let value = parent.child("w", child)?.attribute(Some("w"), "val")?;
    allowed.contains(&value).then(|| value.to_owned())
}

fn attribute_owned(element: &XmlElement, name: &str) -> Option<String> {
    element.attribute(Some("w"), name).map(str::to_owned)
}

fn attribute_nonempty(element: &XmlElement, name: &str) -> Option<String> {
    element
        .attribute(Some("w"), name)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

fn matches_ci(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

fn merge_table_borders(target: Option<&Borders>, source: &Borders) -> Borders {
    let mut result = target.cloned().unwrap_or_default();
    merge_border_slot(
        &mut result.top,
        target.and_then(|value| value.top.as_ref()),
        &source.top,
    );
    merge_border_slot(
        &mut result.bottom,
        target.and_then(|value| value.bottom.as_ref()),
        &source.bottom,
    );
    merge_border_slot(
        &mut result.left,
        target.and_then(|value| value.left.as_ref()),
        &source.left,
    );
    merge_border_slot(
        &mut result.right,
        target.and_then(|value| value.right.as_ref()),
        &source.right,
    );
    merge_border_slot(
        &mut result.inside_h,
        target.and_then(|value| value.inside_h.as_ref()),
        &source.inside_h,
    );
    merge_border_slot(
        &mut result.inside_v,
        target.and_then(|value| value.inside_v.as_ref()),
        &source.inside_v,
    );
    merge_border_slot(
        &mut result.between,
        target.and_then(|value| value.between.as_ref()),
        &source.between,
    );
    merge_border_slot(
        &mut result.bar,
        target.and_then(|value| value.bar.as_ref()),
        &source.bar,
    );
    merge_border_slot(
        &mut result.start,
        target.and_then(|value| value.start.as_ref()),
        &source.start,
    );
    merge_border_slot(
        &mut result.end,
        target.and_then(|value| value.end.as_ref()),
        &source.end,
    );
    merge_border_slot(
        &mut result.tl2br,
        target.and_then(|value| value.tl2br.as_ref()),
        &source.tl2br,
    );
    merge_border_slot(
        &mut result.tr2bl,
        target.and_then(|value| value.tr2bl.as_ref()),
        &source.tr2bl,
    );
    result
}

fn merge_border_slot(
    output: &mut Option<BorderSpec>,
    target: Option<&BorderSpec>,
    source: &Option<BorderSpec>,
) {
    if let Some(source) = source {
        let mut value = target.cloned().unwrap_or_else(|| source.clone());
        value.style.clone_from(&source.style);
        if let Some(color) = &source.color {
            value.color = Some(merge_color_deep(
                target.and_then(|value| value.color.as_ref()),
                color,
            ));
        }
        overlay(&mut value.size, &source.size);
        overlay(&mut value.space, &source.space);
        overlay(&mut value.shadow, &source.shadow);
        overlay(&mut value.frame, &source.frame);
        *output = Some(value);
    }
}

fn merge_shading_deep(
    target: Option<&ShadingProperties>,
    source: &ShadingProperties,
) -> ShadingProperties {
    let mut value = target.cloned().unwrap_or_default();
    if let Some(color) = &source.color {
        value.color = Some(merge_color_deep(
            target.and_then(|value| value.color.as_ref()),
            color,
        ));
    }
    if let Some(fill) = &source.fill {
        value.fill = Some(merge_color_deep(
            target.and_then(|value| value.fill.as_ref()),
            fill,
        ));
    }
    overlay(&mut value.pattern, &source.pattern);
    value
}

fn merge_color_deep(target: Option<&ColorValue>, source: &ColorValue) -> ColorValue {
    let mut value = target.cloned().unwrap_or_default();
    overlay(&mut value.rgb, &source.rgb);
    overlay(&mut value.theme_color, &source.theme_color);
    overlay(&mut value.theme_tint, &source.theme_tint);
    overlay(&mut value.theme_shade, &source.theme_shade);
    overlay(&mut value.auto, &source.auto);
    value
}

fn overlay<T: Clone>(target: &mut Option<T>, source: &Option<T>) {
    if source.is_some() {
        target.clone_from(source);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(
            xml.as_bytes(),
            "formatting.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    #[test]
    fn parses_full_style_property_bags_with_incumbent_defaults() {
        let style = root(
            r#"<w:style><w:pPr><w:spacing w:before="120" w:beforeAutospacing="false"/><w:ind w:firstLine="20" w:hanging="30"/><w:numPr><w:ilvl w:val="2"/></w:numPr><w:rPr><w:b w:val="off"/><w:lang w:val="en-US"/></w:rPr></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi" w:cstheme="majorBidi"/><w:u w:val="double"/><w:color w:val="auto"/></w:rPr><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="off"/></w:tblPr><w:tcPr><w:vMerge/></w:tcPr></w:style>"#,
        );
        let paragraph = parse_paragraph_properties(style.child("w", "pPr"), None).unwrap();
        assert_eq!(paragraph.indent_first_line, Some(-30.0));
        assert_eq!(paragraph.before_autospacing, Some(false));
        assert_eq!(paragraph.num_pr.unwrap().ilvl, Some(2.0));
        let run = parse_run_properties(style.child("w", "rPr"), None).unwrap();
        assert_eq!(
            run.font_family.unwrap().cs_theme.as_deref(),
            Some("majorBidi")
        );
        assert_eq!(run.color.unwrap().auto, Some(true));
        let table = parse_table_properties(style.child("w", "tblPr")).unwrap();
        assert_eq!(table.width.unwrap().value, 5000.0);
        assert_eq!(table.look.unwrap().first_row, Some(false));
        assert_eq!(
            parse_table_cell_properties(style.child("w", "tcPr"))
                .unwrap()
                .v_merge
                .as_deref(),
            Some("continue")
        );
    }

    #[test]
    fn merge_preserves_auto_color_and_replaces_font_pairs_bug_compatibly() {
        let target = TextFormatting {
            color: Some(ColorValue {
                rgb: Some("FF0000".into()),
                ..ColorValue::default()
            }),
            font_family: Some(FontFamily {
                ascii: Some("Calibri".into()),
                ascii_theme: Some("minorHAnsi".into()),
                ..FontFamily::default()
            }),
            ..TextFormatting::default()
        };
        let source = TextFormatting {
            color: Some(ColorValue {
                auto: Some(true),
                ..ColorValue::default()
            }),
            font_family: Some(FontFamily {
                ascii: Some("Arial".into()),
                ..FontFamily::default()
            }),
            ..TextFormatting::default()
        };
        let merged = merge_text_formatting(Some(&target), Some(&source)).unwrap();
        assert_eq!(merged.color.unwrap().rgb.as_deref(), Some("FF0000"));
        let family = merged.font_family.unwrap();
        assert_eq!(family.ascii.as_deref(), Some("Arial"));
        assert_eq!(family.ascii_theme, None);
    }

    #[test]
    fn garbage_numbers_and_language_tags_are_omitted_without_panics() {
        let run = root(&format!(
            r#"<w:rPr><w:sz w:val="{}"/><w:lang w:val="bad tag"/></w:rPr>"#,
            "9".repeat(10_000)
        ));
        assert_eq!(parse_run_properties(Some(&run), None), None);
    }

    #[test]
    fn preserves_incumbent_empty_bags_created_by_undefined_assignments() {
        let conditional = root(
            r#"<w:tblStylePr><w:pPr><w:tabs/></w:pPr><w:rPr><w:shd/></w:rPr><w:tblPr><w:tblW/></w:tblPr><w:trPr><w:trHeight/></w:trPr><w:tcPr><w:tcBorders/></w:tcPr></w:tblStylePr>"#,
        );
        assert_eq!(
            parse_paragraph_properties(conditional.child("w", "pPr"), None),
            Some(ParagraphFormatting::default())
        );
        assert_eq!(
            parse_run_properties(conditional.child("w", "rPr"), None),
            Some(TextFormatting::default())
        );
        assert_eq!(
            parse_table_properties(conditional.child("w", "tblPr")),
            Some(TableFormatting::default())
        );
        assert_eq!(
            parse_table_row_properties(conditional.child("w", "trPr")),
            Some(TableRowFormatting::default())
        );
        assert_eq!(
            parse_table_cell_properties(conditional.child("w", "tcPr")),
            Some(TableCellFormatting::default())
        );
    }
}
