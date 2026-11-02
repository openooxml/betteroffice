//! text functions beyond the everyday set: code points, currency and thai
//! number words, width conversion, values written out as text, and url
//! encoding.

use xlsx_model::{CellValue, ErrorValue};

use crate::array::evaluate_array;
use crate::eval::{
    EvalContext, as_area, err, evaluate, num, text, to_text, used_height, used_width,
};
use crate::parser::Expr;

use super::engineering::significant;
use super::text::fixed_text;
use super::{nth_int, nth_number, omitted};

/// UNICHAR(number): the character with that code point. zero or a number past
/// the last code point is `#VALUE!`; a lone surrogate half is `#N/A`.
pub(crate) fn unichar(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match nth_number(args, ctx, 0).map(f64::trunc) {
        Ok(n) if !(1.0..=f64::from(0x10_FFFF)).contains(&n) => err(ErrorValue::Value),
        Ok(n) => match char::from_u32(n as u32) {
            Some(c) => text(ctx, c.to_string()),
            None => err(ErrorValue::NA),
        },
        Err(e) => err(e),
    }
}

/// UNICODE(text): the code point of the first character; empty text is
/// `#VALUE!`.
pub(crate) fn unicode(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => match s.chars().next() {
            Some(c) => num(f64::from(u32::from(c))),
            None => err(ErrorValue::Value),
        },
        Err(e) => err(e),
    }
}

/// DOLLAR(number, [decimals]): the number rounded like FIXED and written in
/// the en-US currency format `$#,##0.00_);($#,##0.00)`, negatives in
/// parentheses. a value that rounds to zero shows no sign.
pub(crate) fn dollar(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(x) => x,
        Err(e) => return err(e),
    };
    let decimals = match args.get(1) {
        Some(_) => match nth_int(args, ctx, 1) {
            Ok(d) if d > 127 => return err(ErrorValue::Value),
            Ok(d) => d.max(-400) as i32,
            Err(e) => return err(e),
        },
        None => 2,
    };
    let amount = fixed_text(x.abs(), decimals, true);
    let negative = x < 0.0 && amount.bytes().any(|b| matches!(b, b'1'..=b'9'));
    if negative {
        text(ctx, format!("(${amount})"))
    } else {
        text(ctx, format!("${amount}"))
    }
}

/// VALUETOTEXT(value, [format]): a value as text. format 0 (concise) writes
/// text as it is; format 1 (strict) quotes it, doubling inner quotes. errors
/// become their text instead of propagating.
pub(crate) fn valuetotext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let strict = match format_flag(args, ctx) {
        Ok(strict) => strict,
        Err(e) => return err(e),
    };
    text(ctx, written(&evaluate(&args[0], ctx), strict))
}

/// ARRAYTOTEXT(array, [format]): every value of an array as text. format 0
/// joins them row by row with `, `; format 1 writes an array constant,
/// `{a,b;c,d}`, with text quoted. numbers keep 15 significant digits, and an
/// error the argument computes, rather than reads from a cell, is the result.
pub(crate) fn arraytotext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let strict = match format_flag(args, ctx) {
        Ok(strict) => strict,
        Err(e) => return err(e),
    };
    let value = evaluate_array(&args[0], ctx);
    if let crate::array::Value::Scalar(CellValue::Error { value }) = value
        && as_area(&args[0], ctx).is_none()
    {
        return err(value);
    }
    let array = value.into_array();
    let (rows, cols) = (array.rows(), array.cols());
    let cells = array.cells();
    let mut out = String::new();
    if strict {
        out.push('{');
    }
    for row in 0..rows {
        for col in 0..cols {
            if row + col > 0 {
                out.push_str(match (strict, col) {
                    (false, _) => ", ",
                    (true, 0) => ";",
                    (true, _) => ",",
                });
            }
            out.push_str(&written(&cells[row * cols + col], strict));
            if out.len() > 4 * crate::eval::MAX_CELL_TEXT_CHARS {
                return err(ErrorValue::Value);
            }
        }
    }
    if strict {
        out.push('}');
    }
    text(ctx, out)
}

/// the format argument of VALUETOTEXT and ARRAYTOTEXT: 0 or omitted is
/// concise, 1 strict, anything else `#VALUE!`.
fn format_flag(args: &[Expr], ctx: &EvalContext<'_>) -> Result<bool, ErrorValue> {
    match args.get(1) {
        None => Ok(false),
        Some(arg) if omitted(arg) => Ok(false),
        Some(_) => match nth_int(args, ctx, 1)? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(ErrorValue::Value),
        },
    }
}

fn written(value: &CellValue, strict: bool) -> String {
    match value {
        CellValue::Error { value } => value.as_str().to_string(),
        CellValue::Number { value } => significant(*value),
        CellValue::Text { value } if strict => format!("\"{}\"", value.replace('"', "\"\"")),
        other => to_text(other).unwrap_or_default(),
    }
}

/// ASC(text): full-width (double-byte) letters, digits, punctuation and
/// katakana become their half-width forms; a voiced katakana splits into
/// its base and a half-width sound mark.
pub(crate) fn asc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => {
            let mut out = String::with_capacity(s.len());
            for c in s.chars() {
                half_width(c, &mut out);
            }
            text(ctx, out)
        }
        Err(e) => err(e),
    }
}

/// JIS(text), and DBCS under its other name: half-width (single-byte)
/// characters become full-width; a half-width katakana followed by a sound
/// mark joins into one voiced full-width katakana.
pub(crate) fn jis(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => text(ctx, full_width(&s)),
        Err(e) => err(e),
    }
}

/// half-width katakana U+FF61..=U+FF9F in order, as their full-width forms.
const KATAKANA: [char; 63] = [
    '\u{3002}', '\u{300C}', '\u{300D}', '\u{3001}', '\u{30FB}', '\u{30F2}', '\u{30A1}', '\u{30A3}',
    '\u{30A5}', '\u{30A7}', '\u{30A9}', '\u{30E3}', '\u{30E5}', '\u{30E7}', '\u{30C3}', '\u{30FC}',
    '\u{30A2}', '\u{30A4}', '\u{30A6}', '\u{30A8}', '\u{30AA}', '\u{30AB}', '\u{30AD}', '\u{30AF}',
    '\u{30B1}', '\u{30B3}', '\u{30B5}', '\u{30B7}', '\u{30B9}', '\u{30BB}', '\u{30BD}', '\u{30BF}',
    '\u{30C1}', '\u{30C4}', '\u{30C6}', '\u{30C8}', '\u{30CA}', '\u{30CB}', '\u{30CC}', '\u{30CD}',
    '\u{30CE}', '\u{30CF}', '\u{30D2}', '\u{30D5}', '\u{30D8}', '\u{30DB}', '\u{30DE}', '\u{30DF}',
    '\u{30E0}', '\u{30E1}', '\u{30E2}', '\u{30E4}', '\u{30E6}', '\u{30E8}', '\u{30E9}', '\u{30EA}',
    '\u{30EB}', '\u{30EC}', '\u{30ED}', '\u{30EF}', '\u{30F3}', '\u{309B}', '\u{309C}',
];

/// latin-1 and won signs with full-width forms U+FFE0..=U+FFE6.
const SIGNS: [char; 7] = [
    '\u{A2}', '\u{A3}', '\u{AC}', '\u{AF}', '\u{A6}', '\u{A5}', '\u{20A9}',
];

const VOICED_MARK: char = '\u{FF9E}';
const SEMI_VOICED_MARK: char = '\u{FF9F}';

fn half_katakana(full: char) -> Option<char> {
    KATAKANA
        .iter()
        .position(|&k| k == full)
        .and_then(|i| char::from_u32(0xFF61 + i as u32))
}

/// full-width katakana whose voiced form is the next code point.
const VOICED_BASES: [char; 20] = [
    '\u{30AB}', '\u{30AD}', '\u{30AF}', '\u{30B1}', '\u{30B3}', '\u{30B5}', '\u{30B7}', '\u{30B9}',
    '\u{30BB}', '\u{30BD}', '\u{30BF}', '\u{30C1}', '\u{30C4}', '\u{30C6}', '\u{30C8}', '\u{30CF}',
    '\u{30D2}', '\u{30D5}', '\u{30D8}', '\u{30DB}',
];

fn voices(base: char) -> bool {
    VOICED_BASES.contains(&base)
}

/// the ha row, whose semi-voiced form is two code points on.
fn semi_voices(base: char) -> bool {
    matches!(
        base,
        '\u{30CF}' | '\u{30D2}' | '\u{30D5}' | '\u{30D8}' | '\u{30DB}'
    )
}

fn half_width(c: char, out: &mut String) {
    let code = u32::from(c);
    let mapped = match code {
        0xFF01..=0xFF5E => char::from_u32(code - 0xFEE0),
        0x3000 => Some(' '),
        0xFFE0..=0xFFE6 => Some(SIGNS[(code - 0xFFE0) as usize]),
        _ => None,
    };
    if let Some(m) = mapped {
        out.push(m);
        return;
    }
    if let Some(m) = half_katakana(c) {
        out.push(m);
        return;
    }
    let previous = |back: u32| code.checked_sub(back).and_then(char::from_u32);
    let split = match c {
        '\u{30F4}' => Some(('\u{30A6}', VOICED_MARK)),
        _ => match (previous(1), previous(2)) {
            (Some(base), _) if voices(base) => Some((base, VOICED_MARK)),
            (_, Some(base)) if semi_voices(base) => Some((base, SEMI_VOICED_MARK)),
            _ => None,
        },
    };
    match split.and_then(|(base, mark)| half_katakana(base).map(|half| (half, mark))) {
        Some((half, mark)) => {
            out.push(half);
            out.push(mark);
        }
        None => out.push(c),
    }
}

fn full_width(s: &str) -> String {
    let mut out = String::with_capacity(s.len() * 3);
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        let code = u32::from(c);
        let full = match code {
            0x21..=0x7E => char::from_u32(code + 0xFEE0).unwrap_or(c),
            0x20 => '\u{3000}',
            0xFF61..=0xFF9F => KATAKANA[(code - 0xFF61) as usize],
            _ => match SIGNS.iter().position(|&sign| sign == c) {
                Some(i) => char::from_u32(0xFFE0 + i as u32).unwrap_or(c),
                None => c,
            },
        };
        let joined = match chars.peek() {
            Some(&VOICED_MARK) if full == '\u{30A6}' => Some('\u{30F4}'),
            Some(&VOICED_MARK) if voices(full) => char::from_u32(u32::from(full) + 1),
            Some(&SEMI_VOICED_MARK) if semi_voices(full) => char::from_u32(u32::from(full) + 2),
            _ => None,
        };
        match joined {
            Some(voiced) if code >= 0xFF61 => {
                chars.next();
                out.push(voiced);
            }
            _ => out.push(full),
        }
    }
    out
}

const THAI_DIGITS: [&str; 10] = [
    "ศูนย์",
    "หนึ่ง",
    "สอง",
    "สาม",
    "สี่",
    "ห้า",
    "หก",
    "เจ็ด",
    "แปด",
    "เก้า",
];

/// BAHTTEXT(number): the amount in thai words, rounded to satang: baht
/// (บาท) then satang (สตางค์), or ถ้วน for a whole amount; ลบ before a
/// negative one. every six digits close with ล้าน.
pub(crate) fn bahttext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(x) => x,
        Err(e) => return err(e),
    };
    let rounded = fixed_text(x.abs(), 2, false);
    let (whole, fraction) = rounded.split_once('.').unwrap_or((&rounded, "00"));
    let satang: u32 = fraction.parse().unwrap_or(0);
    let whole = whole.trim_start_matches('0');
    let mut out = String::new();
    if x < 0.0 && (!whole.is_empty() || satang > 0) {
        out.push_str("ลบ");
    }
    if whole.is_empty() && satang == 0 {
        out.push_str("ศูนย์บาทถ้วน");
        return text(ctx, out);
    }
    if !whole.is_empty() {
        let digits = whole.as_bytes();
        let head = match digits.len() % 6 {
            0 => 6,
            rest => rest,
        };
        let mut start = 0;
        let mut end = head;
        while start < digits.len() {
            if start > 0 {
                out.push_str("ล้าน");
            }
            let group = digits[start..end]
                .iter()
                .fold(0u32, |acc, d| acc * 10 + u32::from(d - b'0'));
            thai_group(group, &mut out);
            start = end;
            end += 6;
        }
        out.push_str("บาท");
    }
    if satang == 0 {
        out.push_str("ถ้วน");
    } else {
        thai_group(satang, &mut out);
        out.push_str("สตางค์");
    }
    text(ctx, out)
}

/// words for 0..=999999; zero writes nothing. a one in the tens place is
/// สิบ alone, a two ยี่สิบ, and a units one after higher digits เอ็ด.
fn thai_group(value: u32, out: &mut String) {
    const PLACES: [(u32, &str); 4] = [
        (100_000, "แสน"),
        (10_000, "หมื่น"),
        (1_000, "พัน"),
        (100, "ร้อย"),
    ];
    let mut rest = value;
    for (place, word) in PLACES {
        let digit = rest / place;
        if digit > 0 {
            out.push_str(THAI_DIGITS[digit as usize]);
            out.push_str(word);
        }
        rest %= place;
    }
    let (tens, ones) = (rest / 10, rest % 10);
    match tens {
        0 => {}
        1 => out.push_str("สิบ"),
        2 => out.push_str("ยี่สิบ"),
        _ => {
            out.push_str(THAI_DIGITS[tens as usize]);
            out.push_str("สิบ");
        }
    }
    match ones {
        0 => {}
        1 if value > 1 => out.push_str("เอ็ด"),
        _ => out.push_str(THAI_DIGITS[ones as usize]),
    }
}

/// PHONETIC(reference): the furigana of a reference's text. the workbook
/// model keeps no phonetic runs, so each cell reads as its own text, which
/// is what excel shows for text typed without furigana. like excel it joins
/// the text constants of a range in row order, skipping numbers, logicals,
/// errors and formula results. a value that is no reference is `#N/A`.
pub(crate) fn phonetic(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let Some(area) = as_area(&args[0], ctx) else {
        return match evaluate(&args[0], ctx) {
            CellValue::Error { value } => err(value),
            _ => err(ErrorValue::NA),
        };
    };
    let rows = area.rows.min(used_height(&area, ctx));
    let cols = area.cols.min(used_width(&area, ctx));
    let mut out = String::new();
    for row in 0..rows {
        for col in 0..cols {
            let value = match area.get(ctx, row, col) {
                Ok(value) => value,
                Err(e) => return err(e),
            };
            let at =
                xlsx_model::CellRef::new(area.start.row + row as u32, area.start.col + col as u32);
            if let CellValue::Text { value } = value
                && ctx.provider.formula(area.sheet, at).is_none()
            {
                out.push_str(&value);
            }
            if out.len() > 4 * crate::eval::MAX_CELL_TEXT_CHARS {
                return err(ErrorValue::Value);
            }
        }
    }
    text(ctx, out)
}

/// ENCODEURL(text): the text's utf-8 bytes percent-encoded, keeping only
/// the unreserved letters, digits and `-._~`.
pub(crate) fn encodeurl(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => {
            let mut out = String::with_capacity(s.len());
            for b in s.bytes() {
                if b.is_ascii_alphanumeric() || matches!(b, b'-' | b'.' | b'_' | b'~') {
                    out.push(char::from(b));
                } else {
                    out.push_str(&format!("%{b:02X}"));
                }
            }
            text(ctx, out)
        }
        Err(e) => err(e),
    }
}

/// `_xlfn._LONGTEXT(text1, [text2], ...)`: the wrapper excel stores around a
/// string literal longer than the 255 characters a formula token holds, split
/// over as many string arguments as it needs; it is their joined text.
pub(crate) fn longtext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        return err(ErrorValue::Value);
    }
    let mut out = String::new();
    let mut chars = 0usize;
    for arg in args {
        let part = match to_text(&evaluate(arg, ctx)) {
            Ok(part) => part,
            Err(e) => return err(e),
        };
        chars = chars.saturating_add(part.chars().count());
        if chars > crate::eval::MAX_CELL_TEXT_CHARS {
            return err(ErrorValue::Value);
        }
        out.push_str(&part);
    }
    text(ctx, out)
}

fn one_text(args: &[Expr], ctx: &EvalContext<'_>) -> Result<String, ErrorValue> {
    match args {
        [arg] => to_text(&evaluate(arg, ctx)),
        _ => Err(ErrorValue::Value),
    }
}
