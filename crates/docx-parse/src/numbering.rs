//! `numbering.xml` definitions, override inheritance, and bounded marker rendering helpers.

use serde::{Deserialize, Serialize};

use crate::formatting::{FontFamily, ParagraphFormatting, TextFormatting};
use crate::scalars::{ColorValue, parse_color_value};
use crate::settings::incumbent_utf8_text_boundary;
use crate::tabs::TabStop;
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberingDefinitions {
    pub abstract_nums: Vec<AbstractNumbering>,
    pub nums: Vec<NumberingInstance>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AbstractNumbering {
    pub abstract_num_id: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub multi_level_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_style_link: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_link: Option<String>,
    pub levels: Vec<ListLevel>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberingInstance {
    pub num_id: f64,
    pub abstract_num_id: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub level_overrides: Option<Vec<LevelOverride>>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LevelOverride {
    pub ilvl: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_override: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lvl: Option<ListLevel>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ListLevel {
    pub ilvl: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<f64>,
    pub num_fmt: String,
    pub lvl_text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lvl_jc: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suffix: Option<String>,
    #[serde(rename = "pPr", skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<ParagraphFormatting>,
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<TextFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lvl_restart: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_lgl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legacy: Option<LegacyLevel>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LegacyLevel {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legacy: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legacy_space: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legacy_indent: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ListRendering {
    pub marker: String,
    pub level: f64,
    pub num_id: f64,
    pub is_bullet: bool,
    pub num_fmt: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_suffix: Option<String>,
    pub level_num_fmts: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub abstract_num_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_override: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct NumberingMap {
    pub definitions: NumberingDefinitions,
}

pub fn parse_numbering(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<NumberingMap, ParseError> {
    let Some(xml) = xml else {
        return Ok(NumberingMap::default());
    };
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(NumberingMap::default());
    }
    let document = parse_xml(xml, part, budget)?;
    let mut definitions = NumberingDefinitions::default();
    let Some(root) = document.root() else {
        return Ok(NumberingMap { definitions });
    };
    for element in root.children_named("w", "abstractNum") {
        budget.charge_leaf_value(part)?;
        if let Some(value) = parse_abstract_numbering(element, part, budget)? {
            definitions.abstract_nums.push(value);
        }
    }
    for element in root.children_named("w", "num") {
        budget.charge_leaf_value(part)?;
        if let Some(value) = parse_numbering_instance(element, part, budget)? {
            definitions.nums.push(value);
        }
    }
    Ok(NumberingMap { definitions })
}

fn parse_abstract_numbering(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<AbstractNumbering>, ParseError> {
    let Some(abstract_num_id) = element.parse_numeric_attribute(Some("w"), "abstractNumId", 1.0)
    else {
        return Ok(None);
    };
    let multi_level_type =
        child_attribute_nonempty(element, "multiLevelType", "val").filter(|value| {
            matches!(
                value.as_str(),
                "hybridMultilevel" | "multilevel" | "singleLevel"
            )
        });
    let mut levels = Vec::new();
    for level in element.children_named("w", "lvl") {
        budget.charge_leaf_value(part)?;
        if let Some(level) = parse_list_level(level) {
            levels.push(level);
        }
    }
    levels.sort_by(|left, right| left.ilvl.total_cmp(&right.ilvl));
    Ok(Some(AbstractNumbering {
        abstract_num_id,
        multi_level_type,
        num_style_link: child_attribute(element, "numStyleLink", "val"),
        style_link: child_attribute(element, "styleLink", "val"),
        levels,
        name: child_attribute(element, "name", "val"),
    }))
}

fn parse_numbering_instance(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<NumberingInstance>, ParseError> {
    let Some(num_id) = element.parse_numeric_attribute(Some("w"), "numId", 1.0) else {
        return Ok(None);
    };
    let Some(abstract_num_id) = numeric_child(element, "abstractNumId", "val") else {
        return Ok(None);
    };
    let override_elements: Vec<_> = element.children_named("w", "lvlOverride").collect();
    let level_overrides = if override_elements.is_empty() {
        None
    } else {
        let mut overrides = Vec::new();
        for element in override_elements {
            budget.charge_leaf_value(part)?;
            let Some(ilvl) = element.parse_numeric_attribute(Some("w"), "ilvl", 1.0) else {
                continue;
            };
            overrides.push(LevelOverride {
                ilvl,
                start_override: numeric_child(element, "startOverride", "val"),
                lvl: element.child("w", "lvl").and_then(parse_list_level),
            });
        }
        Some(overrides)
    };
    Ok(Some(NumberingInstance {
        num_id,
        abstract_num_id,
        level_overrides,
    }))
}

fn parse_list_level(element: &XmlElement) -> Option<ListLevel> {
    let ilvl = element.parse_numeric_attribute(Some("w"), "ilvl", 1.0)?;
    if !(0.0..=8.0).contains(&ilvl) {
        return None;
    }
    let mut level = ListLevel {
        ilvl,
        start: numeric_child(element, "start", "val"),
        num_fmt: "decimal".to_owned(),
        lvl_text: String::new(),
        lvl_jc: child_attribute_nonempty(element, "lvlJc", "val")
            .filter(|value| matches!(value.as_str(), "left" | "center" | "right")),
        suffix: child_attribute_nonempty(element, "suff", "val")
            .filter(|value| matches!(value.as_str(), "tab" | "space" | "nothing")),
        p_pr: element
            .child("w", "pPr")
            .map(parse_level_paragraph_properties),
        r_pr: element.child("w", "rPr").map(parse_level_run_properties),
        lvl_restart: numeric_child(element, "lvlRestart", "val"),
        is_lgl: element
            .child("w", "isLgl")
            .map(|element| element.parse_boolean("w")),
        legacy: element.child("w", "legacy").map(|legacy| LegacyLevel {
            legacy: Some(legacy.parse_boolean("w")),
            legacy_space: legacy.parse_numeric_attribute(Some("w"), "legacySpace", 1.0),
            legacy_indent: legacy.parse_numeric_attribute(Some("w"), "legacyIndent", 1.0),
        }),
    };
    if let Some(num_fmt) = element.child("w", "numFmt") {
        level.num_fmt = resolve_num_fmt(Some(num_fmt)).unwrap_or_else(|| "decimal".to_owned());
    } else if let Some(alternate) = element.child("mc", "AlternateContent") {
        let choice = alternate
            .child("mc", "Choice")
            .and_then(|element| element.child("w", "numFmt"));
        let fallback = alternate
            .child("mc", "Fallback")
            .and_then(|element| element.child("w", "numFmt"));
        if let Some(value) = resolve_num_fmt(choice).or_else(|| resolve_num_fmt(fallback)) {
            level.num_fmt = value;
        }
    }
    if let Some(text) = element.child("w", "lvlText") {
        level.lvl_text = text
            .attribute(Some("w"), "val")
            .unwrap_or_default()
            .to_owned();
    }
    Some(level)
}

fn resolve_num_fmt(element: Option<&XmlElement>) -> Option<String> {
    let element = element?;
    let value = element
        .attribute(Some("w"), "val")
        .filter(|value| !value.is_empty())?;
    if value == "custom" {
        return parse_custom_number_format(element.attribute(Some("w"), "format"));
    }
    is_known_format(value).then(|| value.to_owned())
}

fn is_known_format(value: &str) -> bool {
    matches!(
        value,
        "decimal"
            | "upperRoman"
            | "lowerRoman"
            | "upperLetter"
            | "lowerLetter"
            | "ordinal"
            | "cardinalText"
            | "ordinalText"
            | "hex"
            | "chicago"
            | "bullet"
            | "none"
            | "decimalZero"
            | "ganada"
            | "chosung"
            | "ideographDigital"
            | "japaneseCounting"
            | "aiueo"
            | "iroha"
            | "decimalFullWidth"
            | "decimalHalfWidth"
            | "japaneseLegal"
            | "japaneseDigitalTenThousand"
            | "decimalEnclosedCircle"
            | "decimalFullWidth2"
            | "aiueoFullWidth"
            | "irohaFullWidth"
            | "decimalEnclosedFullstop"
            | "decimalEnclosedParen"
            | "decimalEnclosedCircleChinese"
            | "ideographEnclosedCircle"
            | "ideographTraditional"
            | "ideographZodiac"
            | "ideographZodiacTraditional"
            | "taiwaneseCounting"
            | "ideographLegalTraditional"
            | "taiwaneseCountingThousand"
            | "taiwaneseDigital"
            | "chineseCounting"
            | "chineseLegalSimplified"
            | "chineseCountingThousand"
            | "koreanDigital"
            | "koreanCounting"
            | "koreanLegal"
            | "koreanDigital2"
            | "vietnameseCounting"
            | "russianLower"
            | "russianUpper"
            | "numberInDash"
            | "hebrew1"
            | "hebrew2"
            | "arabicAlpha"
            | "arabicAbjad"
            | "hindiVowels"
            | "hindiConsonants"
            | "hindiNumbers"
            | "hindiCounting"
            | "thaiLetters"
            | "thaiNumbers"
            | "thaiCounting"
    )
}

fn parse_custom_number_format(value: Option<&str>) -> Option<String> {
    let first = value
        .and_then(|value| value.split(',').next())
        .unwrap_or_default()
        .trim();
    if first.len() < 2
        || !first.ends_with('1')
        || !first[..first.len() - 1].bytes().all(|byte| byte == b'0')
    {
        return None;
    }
    match first.len().min(5) {
        2 => Some("decimalZero".to_owned()),
        3 => Some("decimalZero3".to_owned()),
        4 => Some("decimalZero4".to_owned()),
        5 => Some("decimalZero5".to_owned()),
        _ => None,
    }
}

fn parse_level_paragraph_properties(element: &XmlElement) -> ParagraphFormatting {
    let mut value = ParagraphFormatting::default();
    if let Some(indent) = element.child("w", "ind") {
        value.indent_left = indent
            .parse_numeric_attribute(Some("w"), "left", 1.0)
            .or_else(|| indent.parse_numeric_attribute(Some("w"), "start", 1.0));
        value.indent_right = indent
            .parse_numeric_attribute(Some("w"), "right", 1.0)
            .or_else(|| indent.parse_numeric_attribute(Some("w"), "end", 1.0));
        if let Some(hanging) = indent.parse_numeric_attribute(Some("w"), "hanging", 1.0) {
            value.indent_first_line = Some(-hanging);
            value.hanging_indent = Some(true);
        } else {
            value.indent_first_line = indent.parse_numeric_attribute(Some("w"), "firstLine", 1.0);
        }
    }
    if let Some(tabs) = element.child("w", "tabs") {
        value.tabs = Some(
            tabs.children_named("w", "tab")
                .filter_map(|tab| {
                    let position = tab.parse_numeric_attribute(Some("w"), "pos", 1.0)?;
                    let alignment = tab
                        .attribute(Some("w"), "val")
                        .filter(|value| !value.is_empty())?;
                    Some(TabStop {
                        position,
                        alignment: match alignment {
                            "left" | "center" | "right" | "decimal" | "bar" | "clear" | "num" => {
                                alignment
                            }
                            _ => "left",
                        }
                        .to_owned(),
                        leader: tab
                            .attribute(Some("w"), "leader")
                            .filter(|leader| {
                                matches!(
                                    *leader,
                                    "none"
                                        | "dot"
                                        | "hyphen"
                                        | "underscore"
                                        | "heavy"
                                        | "middleDot"
                                )
                            })
                            .map(str::to_owned),
                    })
                })
                .collect(),
        );
    }
    value
}

fn parse_level_run_properties(element: &XmlElement) -> TextFormatting {
    let mut value = TextFormatting::default();
    value.font_family = element.child("w", "rFonts").map(|fonts| FontFamily {
        ascii: attribute_owned(fonts, "ascii"),
        h_ansi: attribute_owned(fonts, "hAnsi"),
        east_asia: attribute_owned(fonts, "eastAsia"),
        cs: attribute_owned(fonts, "cs"),
        ..FontFamily::default()
    });
    value.font_size = numeric_child(element, "sz", "val");
    value.color = element.child("w", "color").and_then(|color| {
        let rgb = color.attribute(Some("w"), "val");
        let theme = color.attribute(Some("w"), "themeColor");
        if rgb == Some("auto") {
            Some(ColorValue {
                auto: Some(true),
                ..ColorValue::default()
            })
        } else if let Some(theme) = theme.filter(|value| !value.is_empty()) {
            Some(parse_color_value(
                None,
                Some(theme),
                color.attribute(Some("w"), "themeTint"),
                color.attribute(Some("w"), "themeShade"),
            ))
        } else {
            rgb.filter(|value| !value.is_empty())
                .map(|rgb| parse_color_value(Some(rgb), None, None, None))
        }
    });
    value.bold = element
        .child("w", "b")
        .map(|element| element.parse_boolean("w"));
    value.italic = element
        .child("w", "i")
        .map(|element| element.parse_boolean("w"));
    value.hidden = element
        .child("w", "vanish")
        .map(|element| element.parse_boolean("w"));
    value
}

impl NumberingMap {
    pub fn get_level(&self, num_id: f64, ilvl: f64) -> Option<ListLevel> {
        let instance = self.get_instance(num_id)?;
        if let Some(level_override) = instance
            .level_overrides
            .as_deref()
            .and_then(|overrides| overrides.iter().find(|value| value.ilvl == ilvl))
        {
            if let Some(level) = &level_override.lvl {
                return Some(level.clone());
            }
            if let Some(abstract_numbering) = self.get_abstract(instance.abstract_num_id)
                && let Some(base) = abstract_numbering
                    .levels
                    .iter()
                    .find(|level| level.ilvl == ilvl)
                && let Some(start) = level_override.start_override
            {
                let mut level = base.clone();
                level.start = Some(start);
                return Some(level);
            }
        }

        let mut abstract_numbering = self.get_abstract(instance.abstract_num_id)?;
        if abstract_numbering
            .num_style_link
            .as_deref()
            .is_some_and(|link| !link.is_empty())
            && abstract_numbering.levels.is_empty()
        {
            let link = abstract_numbering.num_style_link.as_deref().unwrap();
            if let Some(candidate) = self.abstract_map_values().into_iter().find(|candidate| {
                candidate.style_link.as_deref() == Some(link) && !candidate.levels.is_empty()
            }) {
                abstract_numbering = candidate;
            }
        }
        abstract_numbering
            .levels
            .iter()
            .find(|level| level.ilvl == ilvl)
            .cloned()
    }

    pub fn get_abstract(&self, abstract_num_id: f64) -> Option<&AbstractNumbering> {
        self.definitions
            .abstract_nums
            .iter()
            .rev()
            .find(|value| value.abstract_num_id == abstract_num_id)
    }

    pub fn get_instance(&self, num_id: f64) -> Option<&NumberingInstance> {
        self.definitions
            .nums
            .iter()
            .rev()
            .find(|value| value.num_id == num_id)
    }

    pub fn has_numbering(&self, num_id: f64) -> bool {
        self.get_instance(num_id).is_some()
    }

    fn abstract_map_values(&self) -> Vec<&AbstractNumbering> {
        let mut values: Vec<&AbstractNumbering> = Vec::new();
        for value in &self.definitions.abstract_nums {
            if let Some(index) = values
                .iter()
                .position(|candidate| candidate.abstract_num_id == value.abstract_num_id)
            {
                values[index] = value;
            } else {
                values.push(value);
            }
        }
        values
    }
}

pub fn compute_list_rendering(
    num_id: Option<f64>,
    ilvl: Option<f64>,
    numbering: &NumberingMap,
) -> Option<ListRendering> {
    let num_id = num_id?;
    if num_id == 0.0 {
        return None;
    }
    let ilvl = ilvl.unwrap_or(0.0);
    // The incumbent loop is only intended for the schema's 0..8 levels. Guard
    // hostile caller-supplied values before converting them into loop bounds.
    if ilvl.fract() != 0.0 || !(0.0..=8.0).contains(&ilvl) {
        return None;
    }
    let level = numbering.get_level(num_id, ilvl)?;
    let mut level_num_fmts = Vec::with_capacity(ilvl as usize + 1);
    for level_index in 0..=ilvl as usize {
        level_num_fmts.push(
            numbering
                .get_level(num_id, level_index as f64)
                .map(|level| level.num_fmt)
                .unwrap_or_else(|| "decimal".to_owned()),
        );
    }
    let instance = numbering.get_instance(num_id);
    let level_override = instance
        .and_then(|instance| instance.level_overrides.as_deref())
        .and_then(|overrides| overrides.iter().find(|value| value.ilvl == ilvl));
    let marker_font_family = level.r_pr.as_ref().and_then(|formatting| {
        formatting.font_family.as_ref().and_then(|family| {
            family
                .ascii
                .as_deref()
                .filter(|value| !value.is_empty())
                .or_else(|| family.h_ansi.as_deref().filter(|value| !value.is_empty()))
                .map(str::to_owned)
        })
    });
    let marker_font_size = level
        .r_pr
        .as_ref()
        .and_then(|formatting| formatting.font_size)
        .filter(|size| *size != 0.0)
        .map(|size| size / 2.0);
    Some(ListRendering {
        marker: level.lvl_text.clone(),
        level: ilvl,
        num_id,
        is_bullet: level.num_fmt == "bullet",
        num_fmt: level.num_fmt.clone(),
        marker_hidden: level
            .r_pr
            .as_ref()
            .and_then(|formatting| (formatting.hidden == Some(true)).then_some(true)),
        marker_font_family,
        marker_font_size,
        marker_suffix: level.suffix.clone(),
        level_num_fmts,
        abstract_num_id: instance.map(|instance| instance.abstract_num_id),
        start_override: level_override.and_then(|value| value.start_override),
    })
}

pub fn format_number(number: i64, format: &str) -> String {
    match format {
        "decimal" => number.to_string(),
        "decimalZero" => pad_decimal(number, 2),
        "decimalZero3" => pad_decimal(number, 3),
        "decimalZero4" => pad_decimal(number, 4),
        "decimalZero5" => pad_decimal(number, 5),
        "upperRoman" => to_roman(number).to_uppercase(),
        "lowerRoman" => to_roman(number),
        "upperLetter" => to_letter(number).to_uppercase(),
        "lowerLetter" => to_letter(number),
        "ordinal" => to_ordinal(number),
        "bullet" => "•".to_owned(),
        "none" => String::new(),
        "decimalEnclosedParen" => format!("({number})"),
        "numberInDash" => format!("-{number}-"),
        _ => number.to_string(),
    }
}

pub fn pad_decimal(number: i64, width: usize) -> String {
    if number < 0 {
        number.to_string()
    } else {
        format!("{number:0width$}")
    }
}

fn to_roman(number: i64) -> String {
    if !(1..=3999).contains(&number) {
        return number.to_string();
    }
    let mut remaining = number;
    let mut output = String::new();
    for (value, numeral) in [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ] {
        while remaining >= value {
            output.push_str(numeral);
            remaining -= value;
        }
    }
    output
}

fn to_letter(number: i64) -> String {
    if number <= 0 {
        return String::new();
    }
    let mut remaining = number;
    let mut output = String::new();
    while remaining > 0 {
        remaining -= 1;
        output.insert(0, char::from(b'a' + (remaining % 26) as u8));
        remaining /= 26;
    }
    output
}

fn to_ordinal(number: i64) -> String {
    let value = number % 100;
    let suffix = if (11..=13).contains(&value) {
        "th"
    } else {
        match number % 10 {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th",
        }
    };
    format!("{number}{suffix}")
}

pub fn render_list_marker(level_text: &str, counters: &[i64], formats: &[String]) -> String {
    let mut output = level_text.to_owned();
    for index in 0..9 {
        let placeholder = format!("%{}", index + 1);
        if output.contains(&placeholder) {
            let counter = counters.get(index).copied().unwrap_or(1);
            let format = formats.get(index).map(String::as_str).unwrap_or("decimal");
            output = output.replacen(&placeholder, &format_number(counter, format), 1);
        }
    }
    output
}

pub fn get_bullet_character(level: &ListLevel) -> String {
    if !level.lvl_text.is_empty() {
        return level.lvl_text.clone();
    }
    let font = level.r_pr.as_ref().and_then(|formatting| {
        formatting.font_family.as_ref().and_then(|family| {
            family
                .ascii
                .as_deref()
                .filter(|value| !value.is_empty())
                .or_else(|| family.h_ansi.as_deref().filter(|value| !value.is_empty()))
        })
    });
    match font.map(str::to_ascii_lowercase).as_deref() {
        Some("symbol") => "•".to_owned(),
        Some(font) if font.contains("wingding") => "❑".to_owned(),
        _ => "•".to_owned(),
    }
}

pub fn is_bullet_level(level: &ListLevel) -> bool {
    matches!(level.num_fmt.as_str(), "bullet" | "none")
}

fn child_attribute(parent: &XmlElement, child: &str, attribute: &str) -> Option<String> {
    parent
        .child("w", child)?
        .attribute(Some("w"), attribute)
        .map(str::to_owned)
}

fn child_attribute_nonempty(parent: &XmlElement, child: &str, attribute: &str) -> Option<String> {
    parent
        .child("w", child)?
        .attribute(Some("w"), attribute)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

fn numeric_child(parent: &XmlElement, child: &str, attribute: &str) -> Option<f64> {
    parent
        .child("w", child)?
        .parse_numeric_attribute(Some("w"), attribute, 1.0)
}

fn attribute_owned(element: &XmlElement, attribute: &str) -> Option<String> {
    element.attribute(Some("w"), attribute).map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn parse(xml: &str) -> NumberingMap {
        let limits = ParseLimits::default();
        parse_numbering(
            Some(xml.as_bytes()),
            "word/numbering.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
    }

    #[test]
    fn parses_levels_custom_formats_and_all_rendering_fields() {
        let numbering = parse(
            r#"<w:numbering><w:abstractNum w:abstractNumId="7"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><mc:AlternateContent><mc:Choice><w:numFmt w:val="custom" w:format="001, 002, ..."/></mc:Choice><mc:Fallback><w:numFmt w:val="lowerLetter"/></mc:Fallback></mc:AlternateContent><w:lvlText w:val="%1.%2"/><w:suff w:val="space"/><w:pPr><w:ind w:start="720" w:hanging="360"/><w:tabs><w:tab w:pos="720" w:val="mystery"/></w:tabs></w:pPr><w:rPr><w:rFonts w:ascii="Symbol"/><w:sz w:val="24"/><w:vanish/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="12"><w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>"#,
        );
        let level = numbering.get_level(12.0, 1.0).unwrap();
        assert_eq!(level.num_fmt, "decimalZero3");
        assert_eq!(level.start, Some(4.0));
        assert_eq!(level.p_pr.as_ref().unwrap().indent_left, Some(720.0));
        assert_eq!(
            level.p_pr.as_ref().unwrap().tabs.as_ref().unwrap()[0].alignment,
            "left"
        );
        let rendering = compute_list_rendering(Some(12.0), Some(1.0), &numbering).unwrap();
        assert_eq!(rendering.level_num_fmts, ["upperRoman", "decimalZero3"]);
        assert_eq!(rendering.marker_hidden, Some(true));
        assert_eq!(rendering.marker_font_size, Some(12.0));
        assert_eq!(rendering.start_override, Some(4.0));
        assert_eq!(
            render_list_marker(&rendering.marker, &[2, 7], &rendering.level_num_fmts),
            "II.007"
        );
    }

    #[test]
    fn full_override_wins_and_num_style_link_keeps_the_start_override_quirk() {
        let numbering = parse(
            r#"<w:numbering><w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Linked"/><w:lvl w:ilvl="0"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="Linked"/></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride></w:num><w:num w:numId="4"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="x"/></w:lvl></w:lvlOverride></w:num></w:numbering>"#,
        );
        let linked = numbering.get_level(3.0, 0.0).unwrap();
        assert_eq!(linked.num_fmt, "lowerLetter");
        assert_eq!(linked.start, None);
        let overridden = numbering.get_level(4.0, 0.0).unwrap();
        assert_eq!(overridden.num_fmt, "bullet");
        assert_eq!(overridden.start, None);
        assert_eq!(
            compute_list_rendering(Some(4.0), Some(0.0), &numbering)
                .unwrap()
                .start_override,
            Some(8.0)
        );
    }

    #[test]
    fn hostile_levels_and_garbage_attributes_do_not_allocate_or_panic() {
        let numbering = parse(&format!(
            r#"<w:numbering><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="999999999"><w:lvlText w:val="bad"/></w:lvl><w:lvl w:ilvl="{}"><w:start w:val="{}"/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="999999999"/></w:num></w:numbering>"#,
            "9".repeat(10_000),
            "9".repeat(10_000)
        ));
        assert!(numbering.definitions.abstract_nums[0].levels.is_empty());
        assert!(compute_list_rendering(Some(2.0), Some(999_999_999.0), &numbering).is_none());
    }

    #[test]
    fn marker_helpers_pin_first_replacement_and_format_fallbacks() {
        assert_eq!(render_list_marker("%1/%1/%2", &[3], &[]), "3/%1/1");
        assert_eq!(format_number(27, "upperLetter"), "AA");
        assert_eq!(format_number(12, "ordinal"), "12th");
        assert_eq!(format_number(-2, "decimalZero4"), "-2");
    }

    #[test]
    fn unsafe_xml_is_rejected_by_the_shared_core() {
        let limits = ParseLimits::default();
        let error = parse_numbering(
            Some(b"<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><w:numbering>&e;</w:numbering>"),
            "word/numbering.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap_err();
        assert!(matches!(error, ParseError::UnsafeXml { .. }));
    }
}
