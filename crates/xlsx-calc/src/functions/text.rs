//! text functions. positions are 1-based, counted in unicode scalar values
//! (excel counts utf-16 units). FIND is case-sensitive, SEARCH case-insensitive.

use std::borrow::Cow;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{
    EvalContext, MAX_CELL_TEXT_CHARS, as_area, boolean, err, evaluate, num, parse_num, text,
    text_unmetered, to_number, to_text,
};
use crate::parser::Expr;

use xlsx_model::numfmt;

use super::nth_int;

pub(crate) fn len(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => num(s.chars().count() as f64),
        Err(e) => err(e),
    }
}

pub(crate) fn upper(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    map_text(args, ctx, |s| s.to_uppercase())
}

pub(crate) fn lower(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    map_text(args, ctx, |s| s.to_lowercase())
}

/// TRIM: collapse runs of spaces to one and strip the ends (excel trims only
/// the ascii space, u+0020).
/// TEXTBEFORE(text, delimiter, [instance], [match_mode], [match_end],
/// [if_not_found]). a negative instance counts from the end. missing
/// delimiter is `#N/A` unless `if_not_found` is given.
pub(crate) fn textbefore(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    split_at_delimiter(args, ctx, true)
}

/// TEXTAFTER, the same rules taking the remainder instead.
pub(crate) fn textafter(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    split_at_delimiter(args, ctx, false)
}

/// byte spans in `source` that `delimiter` matches, left to right and
/// non-overlapping. the end is returned rather than assumed from the
/// delimiter's own length, because a case-insensitive match can cover a
/// different number of bytes than it.
fn delimiter_spans(source: &str, delimiter: &str, insensitive: bool) -> Vec<(usize, usize)> {
    let span_at = |rest: &str| -> Option<usize> {
        if !insensitive {
            return rest.starts_with(delimiter).then_some(delimiter.len());
        }
        let mut wanted = delimiter.chars().flat_map(char::to_lowercase).peekable();
        let mut taken = 0;
        for character in rest.chars() {
            if wanted.peek().is_none() {
                break;
            }
            for lowered in character.to_lowercase() {
                match wanted.next() {
                    Some(expected) if expected == lowered => {}
                    _ => return None,
                }
            }
            taken += character.len_utf8();
        }
        wanted.peek().is_none().then_some(taken)
    };
    let mut spans = Vec::new();
    let mut at = 0;
    while at < source.len() {
        match span_at(&source[at..]).filter(|length| *length > 0) {
            Some(length) => {
                spans.push((at, at + length));
                at += length;
            }
            None => at += source[at..].chars().next().map_or(1, char::len_utf8),
        }
    }
    spans
}

fn split_at_delimiter(args: &[Expr], ctx: &EvalContext<'_>, before: bool) -> CellValue {
    if args.len() < 2 || args.len() > 6 {
        return err(ErrorValue::Value);
    }
    let source = match nth_text(args, ctx, 0) {
        Ok(t) => t,
        Err(e) => return err(e),
    };
    let delimiter = match nth_text(args, ctx, 1) {
        Ok(d) => d,
        Err(e) => return err(e),
    };
    let instance = match given(args, 2) {
        Some(_) => match nth_int(args, ctx, 2) {
            Ok(0) => return err(ErrorValue::Value),
            Ok(i) => i,
            Err(e) => return err(e),
        },
        None => 1,
    };
    let insensitive = match given(args, 3) {
        Some(_) => match nth_int(args, ctx, 3) {
            Ok(m) => m != 0,
            Err(e) => return err(e),
        },
        None => false,
    };
    let match_end = match given(args, 4) {
        Some(_) => match nth_int(args, ctx, 4) {
            Ok(m) => m != 0,
            Err(e) => return err(e),
        },
        None => false,
    };
    if delimiter.is_empty() {
        return text(ctx, if before { String::new() } else { source });
    }
    // offsets must index `source`, so a case-insensitive search compares in
    // place rather than searching a lowercased copy: unicode case mappings
    // change byte length, and `İ` would slide every later offset.
    let mut spans = delimiter_spans(&source, &delimiter, insensitive);
    if match_end {
        spans.push((source.len(), source.len()));
    }
    let index = if instance > 0 {
        instance as usize - 1
    } else {
        match spans.len().checked_sub(instance.unsigned_abs() as usize) {
            Some(i) => i,
            None => return not_found(args, ctx),
        }
    };
    let Some(&(start, end)) = spans.get(index) else {
        return not_found(args, ctx);
    };
    let out = if before {
        source[..start].to_owned()
    } else {
        source[end..].to_owned()
    };
    text(ctx, out)
}

fn not_found(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match given(args, 5) {
        Some(fallback) => evaluate(fallback, ctx),
        None => err(ErrorValue::NA),
    }
}

/// an argument the call actually supplies; `f(a,b,,,,c)` leaves the skipped
/// ones as blanks, which mean "use the default", not "use zero".
fn given(args: &[Expr], index: usize) -> Option<&Expr> {
    args.get(index)
        .filter(|arg| !crate::functions::omitted(arg))
}

pub(crate) fn trim(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    map_text(args, ctx, |s| {
        s.split(' ')
            .filter(|p| !p.is_empty())
            .collect::<Vec<_>>()
            .join(" ")
    })
}

/// PROPER: capitalize the first letter of each word, lowercase the rest.
pub(crate) fn proper(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    map_text(args, ctx, |s| {
        let mut out = String::with_capacity(s.len());
        let mut prev_alpha = false;
        for ch in s.chars() {
            if ch.is_alphabetic() {
                if prev_alpha {
                    out.extend(ch.to_lowercase());
                } else {
                    out.extend(ch.to_uppercase());
                }
                prev_alpha = true;
            } else {
                out.push(ch);
                prev_alpha = false;
            }
        }
        out
    })
}

/// CLEAN: strip non-printable control characters (below u+0020).
pub(crate) fn clean(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    map_text(args, ctx, |s| {
        s.chars().filter(|c| !c.is_control()).collect()
    })
}

/// LEFT(text, [count]); count defaults to 1.
pub(crate) fn left(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    take_end(args, ctx, true)
}

/// RIGHT(text, [count]); count defaults to 1.
pub(crate) fn right(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    take_end(args, ctx, false)
}

/// MID(text, start, count): 1-based start, `count` characters.
pub(crate) fn mid(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let start = match nth_int(args, ctx, 1) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let count = match nth_int(args, ctx, 2) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    if start < 1 || count < 0 {
        return err(ErrorValue::Value);
    }
    let chars: Vec<char> = s.chars().collect();
    let from = usize::try_from(start - 1)
        .unwrap_or(usize::MAX)
        .min(chars.len());
    let count = usize::try_from(count).unwrap_or(usize::MAX);
    let to = from.saturating_add(count).min(chars.len());
    limited_text(ctx, chars[from..to].iter().collect())
}

/// FIND(find_text, within_text, [start]): case-sensitive; 1-based; not found
/// or an out-of-range start -> #VALUE!.
pub(crate) fn find(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    locate(args, ctx, true)
}

/// SEARCH(find_text, within_text, [start]): case-insensitive.
pub(crate) fn search(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    locate(args, ctx, false)
}

/// SUBSTITUTE(text, old, new, [instance]): replace `old` with `new`; with
/// `instance` only that occurrence (1-based) is replaced.
pub(crate) fn substitute(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 && args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let old = match nth_text(args, ctx, 1) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let new = match nth_text(args, ctx, 2) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    if old.is_empty() {
        return limited_text(ctx, s);
    }
    let instance = if args.len() == 4 {
        match nth_int(args, ctx, 3) {
            Ok(n) if n >= 1 => Some(usize::try_from(n).unwrap_or(usize::MAX)),
            Ok(_) => return err(ErrorValue::Value),
            Err(e) => return err(e),
        }
    } else {
        None
    };
    let mut out = String::new();
    let mut chars = 0;
    let mut rest = s.as_str();
    let mut seen = 0;
    while let Some(pos) = rest.find(&old) {
        seen += 1;
        if !append_limited(&mut out, &rest[..pos], &mut chars) {
            return err(ErrorValue::Value);
        }
        let replacement = if instance.is_none() || instance == Some(seen) {
            &new
        } else {
            &old
        };
        if !append_limited(&mut out, replacement, &mut chars) {
            return err(ErrorValue::Value);
        }
        rest = &rest[pos + old.len()..];
    }
    if !append_limited(&mut out, rest, &mut chars) {
        return err(ErrorValue::Value);
    }
    text(ctx, out)
}

/// REPLACE(old_text, start, num_chars, new_text): positional replacement.
pub(crate) fn replace(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let start = match nth_int(args, ctx, 1) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let count = match nth_int(args, ctx, 2) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let new = match nth_text(args, ctx, 3) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    if start < 1 || count < 0 {
        return err(ErrorValue::Value);
    }
    let chars: Vec<char> = s.chars().collect();
    let from = usize::try_from(start - 1)
        .unwrap_or(usize::MAX)
        .min(chars.len());
    let count = usize::try_from(count).unwrap_or(usize::MAX);
    let to = from.saturating_add(count).min(chars.len());
    let mut out = String::new();
    let mut output_chars = 0;
    let prefix: String = chars[..from].iter().collect();
    let suffix: String = chars[to..].iter().collect();
    if !append_limited(&mut out, &prefix, &mut output_chars)
        || !append_limited(&mut out, &new, &mut output_chars)
        || !append_limited(&mut out, &suffix, &mut output_chars)
    {
        return err(ErrorValue::Value);
    }
    text(ctx, out)
}

/// REPT(text, count): repeat. negative count -> #VALUE!.
pub(crate) fn rept(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    match nth_int(args, ctx, 1) {
        Ok(n) if n < 0 => err(ErrorValue::Value),
        Ok(_) if s.is_empty() => text(ctx, ""),
        Ok(n) => {
            let Ok(count) = usize::try_from(n) else {
                return err(ErrorValue::Value);
            };
            let Some(chars) = s.chars().count().checked_mul(count) else {
                return err(ErrorValue::Value);
            };
            let Some(bytes) = s.len().checked_mul(count) else {
                return err(ErrorValue::Value);
            };
            if chars > MAX_CELL_TEXT_CHARS {
                return err(ErrorValue::Value);
            }
            if !ctx.consume_text(bytes) {
                return err(ErrorValue::Num);
            }
            text_unmetered(s.repeat(count))
        }
        Err(e) => err(e),
    }
}

/// EXACT(a, b): case-sensitive equality.
pub(crate) fn exact(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_text(args, ctx, 0), nth_text(args, ctx, 1)) {
        (Ok(a), Ok(b)) => boolean(a == b),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// T(value): the value if it is text, otherwise an empty string.
pub(crate) fn t(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match evaluate(&args[0], ctx) {
        v @ CellValue::Text { .. } => v,
        CellValue::Error { value } => err(value),
        _ => text(ctx, ""),
    }
}

/// Windows-1252 characters for codes 0x80..=0x9F; the five unassigned codes map
/// to the C1 control of the same value, as Windows does.
const WINDOWS_1252_HIGH: [char; 32] = [
    '\u{20AC}', '\u{81}', '\u{201A}', '\u{192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
    '\u{2C6}', '\u{2030}', '\u{160}', '\u{2039}', '\u{152}', '\u{8D}', '\u{17D}', '\u{8F}',
    '\u{90}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
    '\u{2DC}', '\u{2122}', '\u{161}', '\u{203A}', '\u{153}', '\u{9D}', '\u{17E}', '\u{178}',
];

/// CHAR(number): the Windows-1252 character for a code in 1..=255.
pub(crate) fn char_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match nth_int(args, ctx, 0) {
        Ok(n @ 0x80..=0x9F) => text(ctx, WINDOWS_1252_HIGH[(n - 0x80) as usize].to_string()),
        Ok(n) if (1..=255).contains(&n) => match char::from_u32(n as u32) {
            Some(c) => text(ctx, c.to_string()),
            None => err(ErrorValue::Value),
        },
        Ok(_) => err(ErrorValue::Value),
        Err(e) => err(e),
    }
}

/// CODE(text): the Windows-1252 code of the first character; characters outside
/// that code page give 63 (`?`). empty text -> #VALUE!.
pub(crate) fn code(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => match s.chars().next() {
            Some(c) => num(f64::from(windows_1252_code(c))),
            None => err(ErrorValue::Value),
        },
        Err(e) => err(e),
    }
}

fn windows_1252_code(c: char) -> u32 {
    if let Some(i) = WINDOWS_1252_HIGH.iter().position(|&h| h == c) {
        return 0x80 + i as u32;
    }
    match c as u32 {
        n @ (0..=0x7F | 0xA0..=0xFF) => n,
        _ => 63,
    }
}

/// VALUE(text): parse text to a number; a trailing `%` scales by 1/100.
pub(crate) fn value(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => match parse_numeric_text(&s) {
            Some(n) => num(n),
            None => err(ErrorValue::Value),
        },
        Err(e) => err(e),
    }
}

/// NUMBERVALUE(text, [decimal_sep], [group_sep]): parse with explicit
/// separators (defaults `.` and `,`).
pub(crate) fn numbervalue(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let decimal = separator(args, ctx, 1, '.');
    let group = separator(args, ctx, 2, ',');
    let decimal = match decimal {
        Ok(c) => c,
        Err(e) => return err(e),
    };
    let group = match group {
        Ok(c) => c,
        Err(e) => return err(e),
    };
    let cleaned: String = s
        .chars()
        .filter(|&c| c != group && !c.is_whitespace())
        .map(|c| if c == decimal { '.' } else { c })
        .collect();
    match parse_numeric_text(&cleaned) {
        Some(n) => num(n),
        None => err(ErrorValue::Value),
    }
}

/// TEXT(value, format): format through the §18.8.30–31 code language via
/// `xlsx_model::numfmt`, assuming the 1900 date system.
pub(crate) fn text_fn(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let value = match evaluate(&args[0], ctx) {
        CellValue::Error { value } => return err(value),
        CellValue::Empty => CellValue::Number { value: 0.0 },
        v => v,
    };
    let format = match nth_text(args, ctx, 1) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    // TEXT(x, "") yields an empty string in excel
    if format.is_empty() {
        return text(ctx, "");
    }
    let formatted = numfmt::format_value(&value, &format, ctx.date_system);
    limited_text(ctx, formatted.text)
}

/// FIXED(number, [decimals], [no_commas]): round half away from zero to
/// `decimals` places (2 when omitted, left of the point when negative) and
/// write the result as text, grouping thousands unless `no_commas`. more than
/// 127 decimals is `#VALUE!`.
pub(crate) fn fixed(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    let x = match super::nth_number(args, ctx, 0) {
        Ok(x) => x,
        Err(e) => return err(e),
    };
    let decimals = if args.len() >= 2 {
        match nth_int(args, ctx, 1) {
            Ok(d) if d > 127 => return err(ErrorValue::Value),
            Ok(d) => d.max(-400) as i32,
            Err(e) => return err(e),
        }
    } else {
        2
    };
    let commas = if args.len() == 3 {
        match to_bool_arg(args, ctx, 2) {
            Ok(no_commas) => !no_commas,
            Err(e) => return err(e),
        }
    } else {
        true
    };
    limited_text(ctx, fixed_text(x, decimals, commas))
}

/// `x` rounded at `decimals` and written out. the rounding works on the 15
/// significant digits excel keeps, so 2.675 rounds up to 2.68 as it displays.
fn fixed_text(x: f64, decimals: i32, commas: bool) -> String {
    let scientific = format!("{:.14e}", x.abs());
    let (mantissa, exponent) = scientific
        .split_once('e')
        .unwrap_or((scientific.as_str(), "0"));
    let mut digits: Vec<u8> = mantissa
        .bytes()
        .filter(u8::is_ascii_digit)
        .map(|b| b - b'0')
        .collect();
    // how many of `digits` stand before the decimal point
    let mut point = exponent.parse::<i32>().unwrap_or(0) + 1;
    let keep = point + decimals;
    if keep < 0 {
        digits.clear();
    } else if (keep as usize) < digits.len() {
        let round_up = digits[keep as usize] >= 5;
        digits.truncate(keep as usize);
        if round_up {
            let mut i = digits.len();
            loop {
                if i == 0 {
                    digits.insert(0, 1);
                    point += 1;
                    break;
                }
                i -= 1;
                if digits[i] == 9 {
                    digits[i] = 0;
                } else {
                    digits[i] += 1;
                    break;
                }
            }
        }
    }
    let digit = |i: i32| {
        let d = usize::try_from(i).ok().and_then(|i| digits.get(i)).copied();
        char::from(b'0' + d.unwrap_or(0))
    };
    let whole: String = (0..point).map(digit).collect();
    let whole = match whole.trim_start_matches('0') {
        "" => "0",
        trimmed => trimmed,
    };
    let mut out = String::new();
    if x < 0.0 && digits.iter().any(|&d| d != 0) {
        out.push('-');
    }
    for (i, c) in whole.chars().enumerate() {
        if commas && i > 0 && (whole.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    if decimals > 0 {
        out.push('.');
        out.extend((point..point + decimals).map(digit));
    }
    out
}

/// TEXTJOIN(delimiter, ignore_empty, text1, ...): join, optionally skipping
/// empty values. text arguments may be ranges.
pub(crate) fn textjoin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let delim = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let ignore_empty = match to_bool_arg(args, ctx, 1) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    let mut output = String::new();
    let mut output_chars = 0;
    let mut first = true;
    for arg in &args[2..] {
        let (values, unread) = match as_area(arg, ctx) {
            Some(area) => match area.values_ref(ctx) {
                Ok(v) => {
                    let unread = area.unread(v.len() as u64);
                    (v, unread)
                }
                Err(e) => return err(e),
            },
            None => (vec![Cow::Owned(evaluate(arg, ctx))], 0),
        };
        for v in values {
            let empty = matches!(v.as_ref(), CellValue::Empty)
                || matches!(v.as_ref(), CellValue::Text { value } if value.is_empty());
            if ignore_empty && empty {
                continue;
            }
            match to_text(v.as_ref()) {
                Ok(s) => {
                    if !first && !append_limited(&mut output, &delim, &mut output_chars) {
                        return err(ErrorValue::Value);
                    }
                    if !append_limited(&mut output, &s, &mut output_chars) {
                        return err(ErrorValue::Value);
                    }
                    first = false;
                }
                Err(e) => return err(e),
            }
        }
        // the blanks past a whole-column or whole-row read still join, each
        // one a delimiter wide
        if !ignore_empty && unread > 0 {
            let separators = unread - u64::from(first);
            first = false;
            if !delim.is_empty() {
                for _ in 0..separators {
                    if !append_limited(&mut output, &delim, &mut output_chars) {
                        return err(ErrorValue::Value);
                    }
                }
            }
        }
    }
    text(ctx, output)
}

/// CONCAT / CONCATENATE: join every argument (ranges flattened, row-major).
pub(crate) fn concat(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let mut out = String::new();
    let mut output_chars = 0;
    for arg in args {
        let values = match as_area(arg, ctx) {
            Some(area) => match area.values_ref(ctx) {
                Ok(v) => v,
                Err(e) => return err(e),
            },
            None => vec![Cow::Owned(evaluate(arg, ctx))],
        };
        for v in values {
            match to_text(v.as_ref()) {
                Ok(s) if append_limited(&mut out, &s, &mut output_chars) => {}
                Ok(_) => return err(ErrorValue::Value),
                Err(e) => return err(e),
            }
        }
    }
    text(ctx, out)
}

fn one_text(args: &[Expr], ctx: &EvalContext<'_>) -> Result<String, ErrorValue> {
    if args.len() != 1 {
        return Err(ErrorValue::Value);
    }
    nth_text(args, ctx, 0)
}

fn nth_text(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<String, ErrorValue> {
    let value = to_text(&evaluate(&args[i], ctx))?;
    if value.chars().count() > MAX_CELL_TEXT_CHARS {
        Err(ErrorValue::Value)
    } else {
        Ok(value)
    }
}

fn map_text(args: &[Expr], ctx: &EvalContext<'_>, f: fn(&str) -> String) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => limited_text(ctx, f(&s)),
        Err(e) => err(e),
    }
}

fn take_end(args: &[Expr], ctx: &EvalContext<'_>, from_left: bool) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let s = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let count = if args.len() == 2 {
        match nth_int(args, ctx, 1) {
            Ok(n) if n < 0 => return err(ErrorValue::Value),
            Ok(n) => usize::try_from(n).unwrap_or(usize::MAX),
            Err(e) => return err(e),
        }
    } else {
        1
    };
    let chars: Vec<char> = s.chars().collect();
    let take = count.min(chars.len());
    let slice = if from_left {
        &chars[..take]
    } else {
        &chars[chars.len() - take..]
    };
    limited_text(ctx, slice.iter().collect())
}

fn locate(args: &[Expr], ctx: &EvalContext<'_>, case_sensitive: bool) -> CellValue {
    if args.len() != 2 && args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let needle = match nth_text(args, ctx, 0) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let haystack = match nth_text(args, ctx, 1) {
        Ok(s) => s,
        Err(e) => return err(e),
    };
    let start = if args.len() == 3 {
        match nth_int(args, ctx, 2) {
            Ok(n) if n >= 1 => usize::try_from(n).unwrap_or(usize::MAX),
            Ok(_) => return err(ErrorValue::Value),
            Err(e) => return err(e),
        }
    } else {
        1
    };
    let hay: Vec<char> = haystack.chars().collect();
    if start > hay.len() + 1 {
        return err(ErrorValue::Value);
    }
    let tail: String = hay[start - 1..].iter().collect();
    let offset = if case_sensitive {
        char_index_of(&tail, &needle)
    } else {
        case_insensitive_char_index_of(&tail, &needle)
    };
    match offset {
        Some(off) => num((start + off) as f64),
        None => err(ErrorValue::Value),
    }
}

fn case_insensitive_char_index_of(haystack: &str, needle: &str) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    let folded_index = char_index_of(&haystack.to_lowercase(), &needle.to_lowercase())?;
    haystack
        .chars()
        .enumerate()
        .flat_map(|(index, character)| character.to_lowercase().map(move |_| index))
        .nth(folded_index)
}

/// position of `needle` in `haystack` measured in characters, not bytes.
fn char_index_of(haystack: &str, needle: &str) -> Option<usize> {
    let byte = haystack.find(needle)?;
    Some(haystack[..byte].chars().count())
}

fn separator(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    i: usize,
    default: char,
) -> Result<char, ErrorValue> {
    if i >= args.len() {
        return Ok(default);
    }
    let s = nth_text(args, ctx, i)?;
    s.chars().next().ok_or(ErrorValue::Value)
}

fn to_bool_arg(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<bool, ErrorValue> {
    Ok(to_number(&evaluate(&args[i], ctx))? != 0.0)
}

fn parse_numeric_text(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if let Some(body) = trimmed.strip_suffix('%') {
        return parse_num(body).map(|n| n / 100.0);
    }
    parse_num(trimmed)
}

fn limited_text(ctx: &EvalContext<'_>, value: String) -> CellValue {
    text(ctx, value)
}

fn append_limited(output: &mut String, value: &str, chars: &mut usize) -> bool {
    let Some(next) = chars.checked_add(value.chars().count()) else {
        return false;
    };
    if next > MAX_CELL_TEXT_CHARS {
        return false;
    }
    output.push_str(value);
    *chars = next;
    true
}
