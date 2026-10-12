use serde::Serialize;
use xlsx_model::numfmt::{FormatPart, FormatPartKind};

use crate::Align;

#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct NumberFormatLayout {
    pub left: f32,
    pub width: f32,
    pub parts: Vec<FormatPart>,
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
            let layout = NumberFormatLayout { left: 2.0, width, parts: formatted.layout.clone() };
            let mut runs = Vec::new();
            paint_number_format(&formatted.text, &layout, Align::Right,
                |text| text.chars().map(|c| match c { '(' => 5.0, ')' => 7.0, ' ' => 3.0, _ => 10.0 }).sum(),
                |text, x| runs.push((text.to_string(), x)));
            assert_eq!(runs[0], ("€".into(), 7.0));
            assert_eq!(runs.last().unwrap(), &("12".into(), width - 25.0));
        }
    }
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
    let extra = if has_fill { (layout.width - fixed).max(0.0) } else { 0.0 };
    let width = fixed + extra;
    let mut x = layout.left + match align {
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
