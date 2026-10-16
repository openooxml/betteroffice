use serde::{Deserialize, Deserializer, Serialize};
use xlsx_model::{
    Alignment, Border, BorderEdge, BorderStyle, CellFormat, CellRange, Color, Fill, Font, HAlign,
    NumberFormat, Protection, VAlign,
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
#[serde(from = "CapturedFormatWire", into = "CapturedFormatWire")]
pub struct CapturedFormat {
    pub rows: u32,
    pub columns: u32,
    pub formats: Vec<CellFormat>,
    /// Per format, the facets applying it keeps from the target cell. Empty
    /// when every format sets all of them.
    pub kept: Vec<KeptFacets>,
    pub source_styles: Vec<Option<u32>>,
    pub source_formats: Vec<PreservedCellFormat>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PreservedCellFormat {
    #[serde(flatten)]
    pub format: CellFormat,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_xf: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub derived_from: Option<u32>,
}

impl PreservedCellFormat {
    pub fn from_style(styles: &xlsx_model::Stylesheet, style: Option<u32>) -> Self {
        let (source_xf, derived_from) = style
            .map(|index| styles.format_source(index))
            .unwrap_or((None, None));
        Self {
            format: styles.cell_format(style),
            source_xf,
            derived_from,
        }
    }
}

/// Facets a captured format can leave to the target cell. A capture written
/// before they were modeled carries none of them, so applying it keeps them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct KeptFacets {
    pub protection: bool,
    pub xf_id: bool,
    pub quote_prefix: bool,
    pub pivot_button: bool,
    pub text_rotation: bool,
    pub indent: bool,
    pub relative_indent: bool,
    pub justify_last_line: bool,
    pub reading_order: bool,
}

impl KeptFacets {
    /// Copies each kept facet of `current` into `format`.
    pub fn keep(&self, current: &CellFormat, format: &mut CellFormat) {
        let Self {
            protection,
            xf_id,
            quote_prefix,
            pivot_button,
            text_rotation,
            indent,
            relative_indent,
            justify_last_line,
            reading_order,
        } = *self;
        if protection {
            format.protection = current.protection;
        }
        if xf_id {
            format.xf_id = current.xf_id;
        }
        if quote_prefix {
            format.quote_prefix = current.quote_prefix;
        }
        if pivot_button {
            format.pivot_button = current.pivot_button;
        }
        let (alignment, from) = (&mut format.alignment, &current.alignment);
        if text_rotation {
            alignment.text_rotation = from.text_rotation;
        }
        if indent {
            alignment.indent = from.indent;
        }
        if relative_indent {
            alignment.relative_indent = from.relative_indent;
        }
        if justify_last_line {
            alignment.justify_last_line = from.justify_last_line;
        }
        if reading_order {
            alignment.reading_order = from.reading_order;
        }
    }
}

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CapturedFormatWire {
    rows: u32,
    columns: u32,
    formats: Vec<CapturedCellFormat>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    source_styles: Vec<Option<u32>>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    source_formats: Vec<PreservedCellFormat>,
}

/// A `CellFormat` as captured: every facet it sets is written, and an absent
/// one keeps the target cell's value.
#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CapturedCellFormat {
    font: Font,
    fill: Fill,
    border: Border,
    number_format: NumberFormat,
    alignment: CapturedAlignment,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    protection: Option<Option<Protection>>,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    xf_id: Option<Option<u32>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    quote_prefix: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pivot_button: Option<bool>,
}

#[derive(Serialize, Deserialize)]
struct CapturedAlignment {
    h: Option<HAlign>,
    v: Option<VAlign>,
    wrap_text: bool,
    shrink_to_fit: bool,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    text_rotation: Option<Option<u16>>,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    indent: Option<Option<u32>>,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    relative_indent: Option<Option<i32>>,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    justify_last_line: Option<Option<bool>>,
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    reading_order: Option<Option<u32>>,
}

fn present<'de, D: Deserializer<'de>, T: Deserialize<'de>>(
    deserializer: D,
) -> Result<Option<T>, D::Error> {
    T::deserialize(deserializer).map(Some)
}

fn unless<T>(kept: bool, value: T) -> Option<T> {
    (!kept).then_some(value)
}

impl CapturedCellFormat {
    fn new(format: CellFormat, kept: KeptFacets) -> Self {
        let CellFormat {
            font,
            fill,
            border,
            number_format,
            alignment,
            protection,
            xf_id,
            quote_prefix,
            pivot_button,
        } = format;
        let Alignment {
            h,
            v,
            wrap_text,
            shrink_to_fit,
            text_rotation,
            indent,
            relative_indent,
            justify_last_line,
            reading_order,
        } = alignment;
        Self {
            font,
            fill,
            border,
            number_format,
            alignment: CapturedAlignment {
                h,
                v,
                wrap_text,
                shrink_to_fit,
                text_rotation: unless(kept.text_rotation, text_rotation),
                indent: unless(kept.indent, indent),
                relative_indent: unless(kept.relative_indent, relative_indent),
                justify_last_line: unless(kept.justify_last_line, justify_last_line),
                reading_order: unless(kept.reading_order, reading_order),
            },
            protection: unless(kept.protection, protection),
            xf_id: unless(kept.xf_id, xf_id),
            quote_prefix: unless(kept.quote_prefix, quote_prefix),
            pivot_button: unless(kept.pivot_button, pivot_button),
        }
    }

    fn into_format(self) -> (CellFormat, KeptFacets) {
        let Self {
            font,
            fill,
            border,
            number_format,
            alignment,
            protection,
            xf_id,
            quote_prefix,
            pivot_button,
        } = self;
        let CapturedAlignment {
            h,
            v,
            wrap_text,
            shrink_to_fit,
            text_rotation,
            indent,
            relative_indent,
            justify_last_line,
            reading_order,
        } = alignment;
        let kept = KeptFacets {
            protection: protection.is_none(),
            xf_id: xf_id.is_none(),
            quote_prefix: quote_prefix.is_none(),
            pivot_button: pivot_button.is_none(),
            text_rotation: text_rotation.is_none(),
            indent: indent.is_none(),
            relative_indent: relative_indent.is_none(),
            justify_last_line: justify_last_line.is_none(),
            reading_order: reading_order.is_none(),
        };
        let format = CellFormat {
            font,
            fill,
            border,
            number_format,
            alignment: Alignment {
                h,
                v,
                wrap_text,
                shrink_to_fit,
                text_rotation: text_rotation.flatten(),
                indent: indent.flatten(),
                relative_indent: relative_indent.flatten(),
                justify_last_line: justify_last_line.flatten(),
                reading_order: reading_order.flatten(),
            },
            protection: protection.flatten(),
            xf_id: xf_id.flatten(),
            quote_prefix: quote_prefix.unwrap_or(false),
            pivot_button: pivot_button.unwrap_or(false),
        };
        (format, kept)
    }
}

impl From<CapturedFormatWire> for CapturedFormat {
    fn from(wire: CapturedFormatWire) -> Self {
        let (formats, kept): (Vec<_>, Vec<_>) = wire
            .formats
            .into_iter()
            .map(CapturedCellFormat::into_format)
            .unzip();
        let any_kept = kept.iter().any(|kept| *kept != KeptFacets::default());
        Self {
            rows: wire.rows,
            columns: wire.columns,
            formats,
            kept: if any_kept { kept } else { Vec::new() },
            source_styles: wire.source_styles,
            source_formats: wire.source_formats,
        }
    }
}

impl From<CapturedFormat> for CapturedFormatWire {
    fn from(format: CapturedFormat) -> Self {
        let CapturedFormat {
            rows,
            columns,
            formats,
            kept,
            source_styles,
            source_formats,
        } = format;
        let formats = formats
            .into_iter()
            .enumerate()
            .map(|(index, format)| {
                CapturedCellFormat::new(format, kept.get(index).copied().unwrap_or_default())
            })
            .collect();
        Self {
            rows,
            columns,
            formats,
            source_styles,
            source_formats,
        }
    }
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
