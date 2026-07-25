//! text functions. positions are 1-based, counted in unicode scalar values
//! (excel counts utf-16 units). FIND is case-sensitive, SEARCH case-insensitive.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{
    EvalContext, MAX_CELL_TEXT_CHARS, as_area, boolean, err, evaluate, num, parse_num, text,
    to_number, to_text,
};
use crate::parser::Expr;

use xlsx_model::DateSystem;
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
    limited_text(chars[from..to].iter().collect())
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
        return limited_text(s);
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
    text(out)
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
    text(out)
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
        Ok(_) if s.is_empty() => text(""),
        Ok(n) => {
            let Ok(count) = usize::try_from(n) else {
                return err(ErrorValue::Value);
            };
            let Some(chars) = s.chars().count().checked_mul(count) else {
                return err(ErrorValue::Value);
            };
            if chars > MAX_CELL_TEXT_CHARS || s.len().checked_mul(count).is_none() {
                return err(ErrorValue::Value);
            }
            text(s.repeat(count))
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
        _ => text(""),
    }
}

/// CHAR(number): the character for a code point in 1..=255.
pub(crate) fn char_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match nth_int(args, ctx, 0) {
        Ok(n) if (1..=255).contains(&n) => match char::from_u32(n as u32) {
            Some(c) => text(c.to_string()),
            None => err(ErrorValue::Value),
        },
        Ok(_) => err(ErrorValue::Value),
        Err(e) => err(e),
    }
}

/// CODE(text): code point of the first character. empty text -> #VALUE!.
pub(crate) fn code(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_text(args, ctx) {
        Ok(s) => match s.chars().next() {
            Some(c) => num(c as u32 as f64),
            None => err(ErrorValue::Value),
        },
        Err(e) => err(e),
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
        return text("");
    }
    let formatted = numfmt::format_value(&value, &format, DateSystem::V1900);
    limited_text(formatted.text)
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
        let values = match as_area(arg, ctx) {
            Some(area) => match area.values(ctx) {
                Ok(v) => v,
                Err(e) => return err(e),
            },
            None => vec![evaluate(arg, ctx)],
        };
        for v in values {
            let empty = matches!(v, CellValue::Empty)
                || matches!(&v, CellValue::Text { value } if value.is_empty());
            if ignore_empty && empty {
                continue;
            }
            match to_text(&v) {
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
    }
    text(output)
}

/// CONCAT / CONCATENATE: join every argument (ranges flattened, row-major).
pub(crate) fn concat(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let mut out = String::new();
    let mut output_chars = 0;
    for arg in args {
        let values = match as_area(arg, ctx) {
            Some(area) => match area.values(ctx) {
                Ok(v) => v,
                Err(e) => return err(e),
            },
            None => vec![evaluate(arg, ctx)],
        };
        for v in values {
            match to_text(&v) {
                Ok(s) if append_limited(&mut out, &s, &mut output_chars) => {}
                Ok(_) => return err(ErrorValue::Value),
                Err(e) => return err(e),
            }
        }
    }
    text(out)
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
        Ok(s) => limited_text(f(&s)),
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
    limited_text(slice.iter().collect())
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
    let (needle, tail): (String, String) = if case_sensitive {
        (needle, hay[start - 1..].iter().collect())
    } else {
        (
            needle.to_lowercase(),
            hay[start - 1..].iter().collect::<String>().to_lowercase(),
        )
    };
    match char_index_of(&tail, &needle) {
        Some(off) => num((start + off) as f64),
        None => err(ErrorValue::Value),
    }
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

fn limited_text(value: String) -> CellValue {
    if value.chars().count() > MAX_CELL_TEXT_CHARS {
        err(ErrorValue::Value)
    } else {
        text(value)
    }
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
