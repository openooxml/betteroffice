use serde::Serialize;
use xlsx_model::numfmt::{FormatPart, FormatPartKind};

use crate::Align;

#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct NumberFormatLayout {
    pub left: f32,
    pub width: f32,
    pub parts: Vec<FormatPart>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub general: Option<f64>,
    #[serde(skip_serializing_if = "is_false")]
    pub overflow: bool,
}

fn is_false(value: &bool) -> bool {
    !*value
}

pub fn paint_number_format(
    text: &str,
    layout: &NumberFormatLayout,
    align: Align,
    mut measure: impl FnMut(&str) -> f32,
    mut paint: impl FnMut(&str, f32),
) {
    let mut cursor = 0;
    let mut fixed = 0.0;
    let mut has_fill = false;
    for part in &layout.parts {
        fixed += measure(&text[cursor..part.offset]);
        if part.kind == FormatPartKind::Padding {
            fixed += measure(part.character.encode_utf8(&mut [0; 4]));
            cursor = part.offset + 1;
        } else {
            has_fill = true;
            cursor = part.offset;
        }
    }
    fixed += measure(&text[cursor..]);
    if layout.overflow && fixed > layout.width {
        let fitted = layout
            .general
            .and_then(|value| xlsx_model::numfmt::fit_general(value, layout.width, &mut measure));
        let fitted = fitted.unwrap_or_else(|| {
            let advance = measure("#");
            let count = if advance > 0.0 && layout.width > 0.0 {
                (layout.width / advance).floor().max(1.0) as usize
            } else {
                0
            };
            "#".repeat(count)
        });
        let width = measure(&fitted);
        let x = layout.left
            + match align {
                Align::Left => 0.0,
                Align::Center => (layout.width - width) / 2.0,
                Align::Right => layout.width - width,
            };
        paint(&fitted, x);
        return;
    }
    let extra = if has_fill {
        (layout.width - fixed).max(0.0)
    } else {
        0.0
    };
    let width = fixed + extra;
    let mut x = layout.left
        + match align {
            Align::Left => 0.0,
            Align::Center => (layout.width - width) / 2.0,
            Align::Right => layout.width - width,
        };
    cursor = 0;
    for part in &layout.parts {
        let literal = &text[cursor..part.offset];
        if !literal.is_empty() {
            paint(literal, x);
            x += measure(literal);
        }
        let mut buffer = [0; 4];
        let character = part.character.encode_utf8(&mut buffer);
        if part.kind == FormatPartKind::Padding {
            x += measure(character);
            cursor = part.offset + 1;
        } else {
            let advance = measure(character);
            if advance > 0.0 && extra >= advance {
                let count = (extra / advance).floor() as usize;
                paint(&character.repeat(count), x);
            }
            x += extra;
            cursor = part.offset;
        }
    }
    if cursor < text.len() {
        paint(&text[cursor..], x);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{CellValue, DateSystem};

    #[test]
    fn accounting_fill_uses_measured_padding_at_two_widths() {
        let formatted = xlsx_model::numfmt::format_value(
            &CellValue::Number { value: 12.0 },
            "_(€* #,##0_)",
            DateSystem::V1900,
        );
        for width in [96.0, 156.0] {
            let layout = NumberFormatLayout {
                left: 2.0,
                width,
                parts: formatted.layout.clone(),
                general: None,
                overflow: true,
            };
            let mut runs = Vec::new();
            paint_number_format(
                &formatted.text,
                &layout,
                Align::Right,
                |text| {
                    text.chars()
                        .map(|c| match c {
                            '(' => 5.0,
                            ')' => 7.0,
                            ' ' => 3.0,
                            _ => 10.0,
                        })
                        .sum()
                },
                |text, x| runs.push((text.to_string(), x)),
            );
            assert_eq!(runs[0], ("€".into(), 7.0));
            assert_eq!(runs.last().unwrap(), &("12".into(), width - 25.0));
        }
    }

    #[test]
    fn numbers_and_dates_fit_without_spilling() {
        for (text, general, width, expected) in [
            ("1/1/2020", None, 26.0, "##"),
            ("1234.56789", Some(1234.56789), 76.0, "1234.57"),
            ("123456789", Some(123456789.0), 66.0, "1E+08"),
            ("123456789", Some(123456789.0), 26.0, "##"),
            ("1234.50", None, 26.0, "##"),
        ] {
            let layout = NumberFormatLayout {
                left: 2.0,
                width,
                parts: Vec::new(),
                general,
                overflow: true,
            };
            let mut runs = Vec::new();
            paint_number_format(
                text,
                &layout,
                Align::Right,
                |s| s.len() as f32 * 10.0,
                |text, x| runs.push((text.to_string(), x)),
            );
            assert_eq!(
                runs,
                [(
                    expected.to_string(),
                    2.0 + width - expected.len() as f32 * 10.0
                )]
            );
        }
    }
}
