//! CELL: what excel reports about a referenced cell's place, contents and
//! formatting.

use xlsx_model::styles::{FormatCode, HAlign};
use xlsx_model::{CellRef, CellValue, ErrorValue, SheetId};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, num, text};
use crate::parser::Expr;

use super::omitted;

/// CELL(info_type, [reference]): one fact about the top-left cell of
/// `reference`.
///
/// three answers depend on things a recalculation cannot see, so they are
/// engine gaps: a cell keeps the value excel cached and only an uncached one
/// shows what is computed here. "filename" is excel's saved path,
/// `dir\[book]sheet`, which the workbook model does not carry; uncached it
/// is empty text, excel's answer for a book never saved. "width" counts
/// characters of the default font, uncached `#VALUE!`. without a reference
/// excel reads the cell changed last, so the formula's own cell stands in. an
/// "address" on another sheet is prefixed by the book name excel would add,
/// so it is a gap too and shows the sheet-qualified address.
pub(crate) fn cell(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let info = match evaluate(&args[0], ctx) {
        CellValue::Text { value } => value.to_ascii_lowercase(),
        CellValue::Error { value } => return err(value),
        _ => return err(ErrorValue::Value),
    };
    let (sheet, at) = match args.get(1).filter(|arg| !omitted(arg)) {
        Some(arg) => match as_area(arg, ctx) {
            Some(area) => (area.sheet, area.start),
            None => {
                return match evaluate(arg, ctx) {
                    CellValue::Error { value } => err(value),
                    _ => err(ErrorValue::Value),
                };
            }
        },
        None => match ctx.cell {
            Some(at) => {
                ctx.record_unsupported_function();
                (ctx.sheet, at)
            }
            None => return err(ErrorValue::Value),
        },
    };
    match info.as_str() {
        "address" => address(ctx, sheet, at),
        "col" => num(f64::from(at.col) + 1.0),
        "row" => num(f64::from(at.row) + 1.0),
        "contents" => contents(ctx, sheet, at),
        "type" => {
            let kind = match contents(ctx, sheet, at) {
                CellValue::Empty => "b",
                CellValue::Text { .. } => "l",
                _ => "v",
            };
            text(ctx, kind)
        }
        "prefix" => {
            let prefix = match contents(ctx, sheet, at) {
                CellValue::Text { .. } => label_prefix(ctx, sheet, at),
                _ => "",
            };
            text(ctx, prefix)
        }
        "protect" => num(flag(locked(ctx, sheet, at))),
        "format" => {
            let shape = format_shape(&format_code(ctx, sheet, at));
            let mut code = shape.code;
            if shape.negative_color {
                code.push('-');
            }
            if shape.parentheses {
                code.push_str("()");
            }
            text(ctx, code)
        }
        "color" => num(flag(
            format_shape(&format_code(ctx, sheet, at)).negative_color,
        )),
        "parentheses" => num(flag(format_shape(&format_code(ctx, sheet, at)).parentheses)),
        "filename" => {
            ctx.record_unsupported_function();
            text(ctx, "")
        }
        "width" => {
            ctx.record_unsupported_function();
            err(ErrorValue::Value)
        }
        _ => err(ErrorValue::Value),
    }
}

fn flag(on: bool) -> f64 {
    if on { 1.0 } else { 0.0 }
}

/// the absolute address; on another sheet qualified by its name.
fn address(ctx: &EvalContext<'_>, sheet: SheetId, at: CellRef) -> CellValue {
    let local = format!(
        "${}${}",
        xlsx_model::addr::col_to_letters(at.col),
        at.row + 1
    );
    if sheet == ctx.sheet {
        return text(ctx, local);
    }
    ctx.record_unsupported_function();
    let name = ctx
        .provider
        .workbook()
        .and_then(|book| book.sheet(sheet))
        .map(|found| found.name.clone())
        .unwrap_or_default();
    text(
        ctx,
        format!("{}!{local}", super::lookups::quoted_sheet(&name)),
    )
}

fn contents(ctx: &EvalContext<'_>, sheet: SheetId, at: CellRef) -> CellValue {
    let area = Area {
        sheet,
        start: at,
        rows: 1,
        cols: 1,
    };
    area.get(ctx, 0, 0).unwrap_or_else(err)
}

/// the style a cell shows with: its own, else its row's or column's.
fn resolved<T>(
    ctx: &EvalContext<'_>,
    sheet: SheetId,
    at: CellRef,
    read: impl FnOnce(&xlsx_model::styles::ResolvedFormat<'_>) -> T,
) -> T {
    let book = ctx.provider.workbook();
    let style = book
        .and_then(|book| book.sheet(sheet))
        .and_then(|found| found.effective_style(at));
    match book {
        Some(book) => read(&book.styles.resolved_format(style)),
        None => read(&xlsx_model::styles::Stylesheet::default().resolved_format(None)),
    }
}

/// the label prefix that encodes text alignment: `'` left, `"` right, `^`
/// centred, `\` filled, empty for any other alignment.
fn label_prefix(ctx: &EvalContext<'_>, sheet: SheetId, at: CellRef) -> &'static str {
    resolved(ctx, sheet, at, |format| match format.alignment.h {
        None | Some(HAlign::General) | Some(HAlign::Left) => "'",
        Some(HAlign::Right) => "\"",
        Some(HAlign::Center) | Some(HAlign::CenterContinuous) => "^",
        Some(HAlign::Fill) => "\\",
        Some(HAlign::Justify) | Some(HAlign::Distributed) => "",
    })
}

/// cells are locked unless their style says otherwise.
fn locked(ctx: &EvalContext<'_>, sheet: SheetId, at: CellRef) -> bool {
    resolved(ctx, sheet, at, |format| {
        format.protection.is_none_or(|protection| protection.locked)
    })
}

fn format_code(ctx: &EvalContext<'_>, sheet: SheetId, at: CellRef) -> String {
    resolved(ctx, sheet, at, |format| match format.number_format {
        FormatCode::Custom(code) => code.to_string(),
        FormatCode::Builtin(id) => xlsx_model::numfmt::builtin_format_code(id)
            .unwrap_or("General")
            .to_string(),
    })
}

/// what CELL reports about a number format: microsoft's code for it, and
/// whether negatives show in colour or positives in parentheses.
struct FormatShape {
    code: String,
    negative_color: bool,
    parentheses: bool,
}

/// the tokens of one format section that CELL's codes tell apart.
#[derive(Default)]
struct Section {
    decimals: usize,
    after_point: bool,
    digits: bool,
    thousands: bool,
    percent: bool,
    scientific: bool,
    fraction: bool,
    currency: bool,
    parentheses: bool,
    color: bool,
    general: bool,
    meridiem: bool,
    /// date and time letters in order: `y`, `d`, `h`, `s`, `m` with its run
    /// length, and `n` for an elapsed-minutes bracket.
    clock: Vec<(char, usize)>,
}

impl Section {
    fn has(&self, letter: char) -> bool {
        self.clock.iter().any(|(token, _)| *token == letter)
    }

    /// the longest `m` run that reads as a month: one neither after an hour
    /// nor before seconds.
    fn month(&self) -> usize {
        let clock = &self.clock;
        (0..clock.len())
            .filter(|&k| clock[k].0 == 'm')
            .filter(|&k| k == 0 || clock[k - 1].0 != 'h')
            .filter(|&k| clock.get(k + 1).is_none_or(|next| next.0 != 's'))
            .map(|k| clock[k].1)
            .max()
            .unwrap_or(0)
    }
}

/// map a format code onto microsoft's CELL("format") table: `G` general,
/// `F`/`,`/`C`/`P`/`S` with the decimal count for fixed, thousands,
/// currency, percent and scientific, `D1`..`D9` for the date and time
/// layouts. the negative section's colour and the positive section's
/// parentheses come back separately.
fn format_shape(code: &str) -> FormatShape {
    let sections: Vec<Section> = split_sections(code)
        .iter()
        .map(String::as_str)
        .map(scan)
        .collect();
    let first = &sections[0];
    let negative_color = sections.get(1).unwrap_or(first).color;
    let code = if !first.clock.is_empty() || first.meridiem {
        date_code(first).to_string()
    } else if (first.general && !first.digits) || first.fraction {
        "G".to_string()
    } else if first.percent {
        format!("P{}", first.decimals)
    } else if first.scientific {
        format!("S{}", first.decimals)
    } else if first.currency {
        format!("C{}", first.decimals)
    } else if first.thousands {
        format!(",{}", first.decimals)
    } else if first.digits {
        format!("F{}", first.decimals)
    } else {
        "G".to_string()
    };
    FormatShape {
        code,
        negative_color,
        parentheses: first.parentheses,
    }
}

fn date_code(section: &Section) -> &'static str {
    let month = section.month();
    let (year, day) = (section.has('y'), section.has('d'));
    let time = section.meridiem || section.has('h') || section.has('s') || section.has('n');
    let date = year || day || month > 0;
    match (date, time) {
        (true, true) => "D4",
        (false, _) => match (section.meridiem, section.has('s')) {
            (true, true) => "D6",
            (true, false) => "D7",
            (false, true) => "D8",
            (false, false) => "D9",
        },
        (true, false) => match (year, day, month) {
            (true, true, 3..) => "D1",
            (false, true, 3..) => "D2",
            (true, false, 1..) => "D3",
            (false, true, 1..) => "D5",
            _ => "D4",
        },
    }
}

/// the `;`-separated sections, keeping quoted text and brackets whole.
fn split_sections(code: &str) -> Vec<String> {
    let mut sections = vec![String::new()];
    let mut chars = code.chars();
    while let Some(ch) = chars.next() {
        let current = sections.last_mut().expect("one section");
        match ch {
            ';' => sections.push(String::new()),
            '"' => {
                current.push(ch);
                for inner in chars.by_ref() {
                    current.push(inner);
                    if inner == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                current.push(ch);
                if let Some(next) = chars.next() {
                    current.push(next);
                }
            }
            _ => current.push(ch),
        }
    }
    sections
}

fn scan(section: &str) -> Section {
    let mut out = Section::default();
    let chars: Vec<char> = section.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        match ch {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    out.currency |= is_currency_symbol(chars[i]);
                    i += 1;
                }
            }
            '\\' => {
                i += 1;
                if let Some(&next) = chars.get(i) {
                    out.currency |= is_currency_symbol(next);
                    out.parentheses |= next == '(';
                }
            }
            '_' | '*' => i += 1,
            '[' => {
                let start = i + 1;
                while i < chars.len() && chars[i] != ']' {
                    i += 1;
                }
                let inner: String = chars[start.min(chars.len())..i].iter().collect();
                bracket(&mut out, &inner);
            }
            '0' | '#' | '?' => {
                out.digits = true;
                if out.after_point && !out.scientific {
                    out.decimals += 1;
                }
            }
            '.' => out.after_point = true,
            ',' => {
                let next = chars.get(i + 1).copied();
                if out.digits && matches!(next, Some('0' | '#' | '?')) {
                    out.thousands = true;
                }
            }
            '%' => out.percent = true,
            'E' | 'e' if matches!(chars.get(i + 1), Some('+' | '-')) => {
                out.scientific = true;
                i += 1;
            }
            '/' => out.fraction |= out.digits,
            '(' => out.parentheses = true,
            'G' | 'g' if starts_with_ignore_case(&chars[i..], "general") => {
                out.general = true;
                i += "general".len() - 1;
            }
            'A' | 'a' if starts_with_ignore_case(&chars[i..], "am/pm") => {
                out.meridiem = true;
                i += "am/pm".len() - 1;
            }
            'A' | 'a' if starts_with_ignore_case(&chars[i..], "a/p") => {
                out.meridiem = true;
                i += "a/p".len() - 1;
            }
            'Y' | 'y' | 'D' | 'd' | 'H' | 'h' | 'S' | 's' | 'M' | 'm' => {
                let letter = ch.to_ascii_lowercase();
                let run = chars[i..]
                    .iter()
                    .take_while(|c| c.to_ascii_lowercase() == letter)
                    .count();
                out.clock.push((letter, run));
                i += run - 1;
            }
            _ => out.currency |= is_currency_symbol(ch),
        }
        i += 1;
    }
    out
}

fn bracket(out: &mut Section, inner: &str) {
    if let Some(rest) = inner.strip_prefix('$') {
        let symbol = rest.split('-').next().unwrap_or_default();
        out.currency |= !symbol.is_empty();
        return;
    }
    const COLORS: [&str; 8] = [
        "black", "blue", "cyan", "green", "magenta", "red", "white", "yellow",
    ];
    let lower = inner.to_ascii_lowercase();
    if COLORS.contains(&lower.as_str()) || lower.starts_with("color") {
        out.color = true;
    } else if let Some(letter) = lower
        .chars()
        .next()
        .filter(|c| matches!(c, 'h' | 'm' | 's'))
    {
        out.clock
            .push((if letter == 'm' { 'n' } else { letter }, 1));
    }
}

fn is_currency_symbol(ch: char) -> bool {
    matches!(ch, '$' | '€' | '£' | '¥' | '¤')
}

fn starts_with_ignore_case(chars: &[char], word: &str) -> bool {
    let mut rest = chars.iter();
    word.chars()
        .all(|w| rest.next().is_some_and(|c| c.eq_ignore_ascii_case(&w)))
}

#[cfg(test)]
mod tests {
    use super::format_shape;

    fn reported(code: &str) -> String {
        let shape = format_shape(code);
        let mut out = shape.code;
        if shape.negative_color {
            out.push('-');
        }
        if shape.parentheses {
            out.push_str("()");
        }
        out
    }

    /// microsoft's CELL("format") table.
    #[test]
    fn format_codes_follow_microsofts_table() {
        for (code, want) in [
            ("General", "G"),
            ("0", "F0"),
            ("#,##0", ",0"),
            ("0.00", "F2"),
            ("#,##0.00", ",2"),
            ("$#,##0_);($#,##0)", "C0"),
            ("$#,##0_);[Red]($#,##0)", "C0-"),
            ("$#,##0.00_);($#,##0.00)", "C2"),
            ("$#,##0.00_);[Red]($#,##0.00)", "C2-"),
            ("0%", "P0"),
            ("0.00%", "P2"),
            ("0.00E+00", "S2"),
            ("# ?/?", "G"),
            ("# ??/??", "G"),
            ("m/d/yy", "D4"),
            ("m/d/yy h:mm", "D4"),
            ("mm/dd/yy", "D4"),
            ("d-mmm-yy", "D1"),
            ("dd-mmm-yy", "D1"),
            ("d-mmm", "D2"),
            ("dd-mmm", "D2"),
            ("mmm-yy", "D3"),
            ("mm/dd", "D5"),
            ("h:mm AM/PM", "D7"),
            ("h:mm:ss AM/PM", "D6"),
            ("h:mm", "D9"),
            ("h:mm:ss", "D8"),
        ] {
            assert_eq!(reported(code), want, "{code}");
        }
    }

    #[test]
    fn parentheses_and_colour_come_from_their_sections() {
        assert_eq!(reported("(0.00)"), "F2()");
        assert_eq!(reported("0;[Blue]-0"), "F0-");
        assert_eq!(reported("[$€-407] #,##0.00"), "C2");
        assert_eq!(reported("[$-409]0.0"), "F1");
        assert_eq!(reported("\"$\"#,##0.00"), "C2");
    }
}
