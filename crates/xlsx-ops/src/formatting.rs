use serde::{Deserialize, Serialize};
use xlsx_model::{
    BorderEdge, BorderStyle, CellFormat, CellRange, Color, Fill, HAlign, NumberFormat, VAlign,
};

use crate::apply::OpError;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StyleProperty {
    Bold,
    Italic,
    Strikethrough,
    FontFamily,
    FontSize,
    TextColor,
    FillColor,
    Borders,
    HorizontalAlignment,
    VerticalAlignment,
    TextWrapping,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum HorizontalAlignment {
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum VerticalAlignment {
    Top,
    Middle,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextWrapping {
    Overflow,
    Wrap,
    Clip,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum BorderPreset {
    All,
    Inner,
    Horizontal,
    Vertical,
    Outer,
    Left,
    Top,
    Right,
    Bottom,
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum BorderLineStyle {
    Solid,
    Dashed,
    Dotted,
    Double,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderPatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset: Option<BorderPreset>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<BorderLineStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StylePatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<BorderPatch>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_alignment: Option<HorizontalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<VerticalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_wrapping: Option<TextWrapping>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub clear: Vec<StyleProperty>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum NumberFormatMutation {
    Automatic,
    PlainText,
    Number,
    Percent,
    Scientific,
    Currency,
    Date,
    Time,
    Custom { pattern: String },
    IncreaseDecimal,
    DecreaseDecimal,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CapturedFormat {
    pub rows: u32,
    pub columns: u32,
    pub formats: Vec<CellFormat>,
}

pub(crate) fn patch_cell_format(
    format: &mut CellFormat,
    patch: &StylePatch,
    range: CellRange,
    row: u32,
    col: u32,
) -> Result<(), OpError> {
    for property in &patch.clear {
        match property {
            StyleProperty::Bold => format.font.bold = false,
            StyleProperty::Italic => format.font.italic = false,
            StyleProperty::Strikethrough => format.font.strike = false,
            StyleProperty::FontFamily => format.font.name = None,
            StyleProperty::FontSize => format.font.size_pt = None,
            StyleProperty::TextColor => format.font.color = None,
            StyleProperty::FillColor => format.fill = Fill::None,
            StyleProperty::Borders => format.border = Default::default(),
            StyleProperty::HorizontalAlignment => format.alignment.h = None,
            StyleProperty::VerticalAlignment => format.alignment.v = None,
            StyleProperty::TextWrapping => {
                format.alignment.wrap_text = false;
                format.alignment.shrink_to_fit = false;
            }
        }
    }
    if let Some(value) = patch.bold {
        format.font.bold = value;
    }
    if let Some(value) = patch.italic {
        format.font.italic = value;
    }
    if let Some(value) = patch.strikethrough {
        format.font.strike = value;
    }
    if let Some(value) = &patch.font_family {
        if value.trim().is_empty() {
            return Err(OpError::InvalidStyle(
                "font family must not be empty".into(),
            ));
        }
        format.font.name = Some(value.clone());
    }
    if let Some(value) = patch.font_size {
        if !value.is_finite() || !(1.0..=400.0).contains(&value) {
            return Err(OpError::InvalidStyle(
                "font size must be between 1 and 400".into(),
            ));
        }
        format.font.size_pt = Some(value);
    }
    if let Some(value) = &patch.text_color {
        format.font.color = Some(Color::Rgb(normalize_color(value)?));
    }
    if let Some(value) = &patch.fill_color {
        format.fill = Fill::Solid(Color::Rgb(normalize_color(value)?));
    }
    if let Some(value) = patch.horizontal_alignment {
        format.alignment.h = Some(match value {
            HorizontalAlignment::Left => HAlign::Left,
            HorizontalAlignment::Center => HAlign::Center,
            HorizontalAlignment::Right => HAlign::Right,
        });
    }
    if let Some(value) = patch.vertical_alignment {
        format.alignment.v = Some(match value {
            VerticalAlignment::Top => VAlign::Top,
            VerticalAlignment::Middle => VAlign::Center,
            VerticalAlignment::Bottom => VAlign::Bottom,
        });
    }
    if let Some(value) = patch.text_wrapping {
        format.alignment.wrap_text = value == TextWrapping::Wrap;
        format.alignment.shrink_to_fit = value == TextWrapping::Clip;
    }
    if let Some(border) = &patch.border {
        patch_border(format, border, range, row, col)?;
    }
    Ok(())
}

pub(crate) fn mutate_number_format(format: &mut CellFormat, mutation: &NumberFormatMutation) {
    format.number_format = match mutation {
        NumberFormatMutation::Automatic => NumberFormat::Builtin { id: 0 },
        NumberFormatMutation::PlainText => NumberFormat::Builtin { id: 49 },
        NumberFormatMutation::Number => NumberFormat::Builtin { id: 4 },
        NumberFormatMutation::Percent => NumberFormat::Builtin { id: 10 },
        NumberFormatMutation::Scientific => NumberFormat::Builtin { id: 11 },
        NumberFormatMutation::Currency => NumberFormat::Custom {
            pattern: "$#,##0.00".into(),
        },
        NumberFormatMutation::Date => NumberFormat::Builtin { id: 14 },
        NumberFormatMutation::Time => NumberFormat::Builtin { id: 20 },
        NumberFormatMutation::Custom { pattern } => NumberFormat::Custom {
            pattern: pattern.clone(),
        },
        NumberFormatMutation::IncreaseDecimal => adjust_decimals(&format.number_format, 1),
        NumberFormatMutation::DecreaseDecimal => adjust_decimals(&format.number_format, -1),
    };
}

fn patch_border(
    format: &mut CellFormat,
    patch: &BorderPatch,
    range: CellRange,
    row: u32,
    col: u32,
) -> Result<(), OpError> {
    let color = patch.color.as_deref().map(normalize_color).transpose()?;
    if patch.preset == Some(BorderPreset::None) {
        format.border = Default::default();
        return Ok(());
    }
    if patch.preset.is_none() {
        for edge in border_edges_mut(&mut format.border).into_iter().flatten() {
            if let Some(style) = patch.style {
                edge.style = border_style(style);
            }
            if let Some(color) = &color {
                edge.color = Some(Color::Rgb(color.clone()));
            }
        }
        return Ok(());
    }
    let preset = patch.preset.expect("preset checked");
    let edge = || BorderEdge {
        style: patch.style.map(border_style).unwrap_or(BorderStyle::Thin),
        color: color.clone().map(Color::Rgb),
    };
    let top = row == range.start.row;
    let bottom = row == range.end.row;
    let left = col == range.start.col;
    let right = col == range.end.col;
    let set_top = matches!(preset, BorderPreset::All)
        || matches!(preset, BorderPreset::Inner | BorderPreset::Horizontal) && !top
        || matches!(preset, BorderPreset::Outer | BorderPreset::Top) && top;
    let set_bottom = matches!(preset, BorderPreset::All)
        || matches!(preset, BorderPreset::Inner | BorderPreset::Horizontal) && !bottom
        || matches!(preset, BorderPreset::Outer | BorderPreset::Bottom) && bottom;
    let set_left = matches!(preset, BorderPreset::All)
        || matches!(preset, BorderPreset::Inner | BorderPreset::Vertical) && !left
        || matches!(preset, BorderPreset::Outer | BorderPreset::Left) && left;
    let set_right = matches!(preset, BorderPreset::All)
        || matches!(preset, BorderPreset::Inner | BorderPreset::Vertical) && !right
        || matches!(preset, BorderPreset::Outer | BorderPreset::Right) && right;
    if set_top {
        format.border.top = Some(edge());
    }
    if set_bottom {
        format.border.bottom = Some(edge());
    }
    if set_left {
        format.border.left = Some(edge());
    }
    if set_right {
        format.border.right = Some(edge());
    }
    Ok(())
}

fn border_edges_mut(border: &mut xlsx_model::Border) -> [Option<&mut BorderEdge>; 4] {
    [
        border.left.as_mut(),
        border.top.as_mut(),
        border.right.as_mut(),
        border.bottom.as_mut(),
    ]
}

fn border_style(style: BorderLineStyle) -> BorderStyle {
    match style {
        BorderLineStyle::Solid => BorderStyle::Thin,
        BorderLineStyle::Dashed => BorderStyle::Dashed,
        BorderLineStyle::Dotted => BorderStyle::Dotted,
        BorderLineStyle::Double => BorderStyle::Double,
    }
}

fn normalize_color(value: &str) -> Result<String, OpError> {
    let hex = value.strip_prefix('#').unwrap_or(value);
    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(OpError::InvalidStyle(format!(
            "invalid color {value:?}; expected #rrggbb"
        )));
    }
    Ok(format!("#{}", hex.to_ascii_lowercase()))
}

fn adjust_decimals(format: &NumberFormat, delta: i8) -> NumberFormat {
    let code = match format {
        NumberFormat::Builtin { id } => xlsx_model::numfmt::builtin_format_code(*id)
            .unwrap_or("General")
            .to_string(),
        NumberFormat::Custom { pattern } => pattern.clone(),
    };
    if is_date_or_time(&code) {
        return format.clone();
    }
    NumberFormat::Custom {
        pattern: adjust_pattern_decimals(&code, delta),
    }
}

fn is_date_or_time(code: &str) -> bool {
    let lower = code.to_ascii_lowercase();
    lower.contains('y')
        || lower.contains('d')
        || lower.contains("h:")
        || lower.contains(":m")
        || lower.contains(":s")
}

fn adjust_pattern_decimals(code: &str, delta: i8) -> String {
    let mut chars = code.chars().collect::<Vec<_>>();
    if code.eq_ignore_ascii_case("general") || code == "@" {
        return if delta > 0 { "0.0" } else { "0" }.into();
    }
    let section_end = chars
        .iter()
        .position(|char| *char == ';')
        .unwrap_or(chars.len());
    let dot = chars[..section_end].iter().position(|char| *char == '.');
    if delta > 0 {
        match dot {
            Some(dot) => {
                let end = chars[dot + 1..section_end]
                    .iter()
                    .position(|char| !matches!(char, '0' | '#' | '?'))
                    .map(|offset| dot + 1 + offset)
                    .unwrap_or(section_end);
                chars.insert(end, '0');
            }
            None => {
                let end = chars[..section_end]
                    .iter()
                    .rposition(|char| matches!(char, '0' | '#' | '?'))
                    .map(|index| index + 1)
                    .unwrap_or(section_end);
                chars.insert(end, '.');
                chars.insert(end + 1, '0');
            }
        }
    } else if let Some(dot) = dot {
        let last = chars[dot + 1..section_end]
            .iter()
            .rposition(|char| matches!(char, '0' | '#' | '?'))
            .map(|offset| dot + 1 + offset);
        if let Some(last) = last {
            chars.remove(last);
            if last == dot + 1 {
                chars.remove(dot);
            }
        }
    }
    chars.into_iter().collect()
}
