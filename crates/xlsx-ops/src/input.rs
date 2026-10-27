//! turning raw editor input into a `CellState`, classified the way excel does:
//! leading `=` is a formula, leading `'` forces text, numbers/booleans coerce.

use xlsx_calc::{EvalContext, evaluate, parse_formula};
use xlsx_model::{CellProvider, CellValue, DateSystem, FormatCode, SheetId};

use crate::op::CellState;

/// the classification of a raw editor string, before it becomes a `CellState`.
#[derive(Debug, Clone, PartialEq)]
pub enum ParsedInput {
    /// empty input — clears the cell back to default.
    Clear,
    /// a formula, carrying the text with the leading `=` stripped.
    Formula(String),
    /// literal text (either forced with a leading `'`, or a plain string).
    Text(String),
    Number(f64),
    Bool(bool),
}

/// classify a raw editor string the way excel does on cell entry.
pub fn parse_input(input: &str) -> ParsedInput {
    parse_input_with_context(input, FormatCode::Builtin(0), DateSystem::V1900, None).0
}

/// Classify input and suggest a built-in format for a General cell.
pub fn parse_input_with_context(
    input: &str,
    format: FormatCode<'_>,
    date_system: DateSystem,
    now_serial: Option<f64>,
) -> (ParsedInput, Option<u16>) {
    if !input.is_empty() && matches!(format, FormatCode::Builtin(49) | FormatCode::Custom("@")) {
        return (
            ParsedInput::Text(input.strip_prefix('\'').unwrap_or(input).into()),
            None,
        );
    }
    let parsed = parse_literal(input);
    if !matches!(parsed, ParsedInput::Text(_)) || input.starts_with('\'') {
        return (parsed, None);
    }
    let trimmed = input.trim();
    let coerced = parse_accounting_negative(trimmed).or_else(|| {
        parse_typed_number(trimmed)
            .or_else(|| parse_timestamp(trimmed, day_first(&format), date_system, now_serial))
            .map(|(value, id)| (value, Some(id)))
    });
    let general = match format {
        FormatCode::Builtin(id) => id == 0,
        FormatCode::Custom(code) => code.eq_ignore_ascii_case("general"),
    };
    match coerced {
        Some((value, id)) => (ParsedInput::Number(value), id.filter(|_| general)),
        None => (parsed, None),
    }
}

fn parse_accounting_negative(input: &str) -> Option<(f64, Option<u16>)> {
    let body = input.strip_prefix('(')?.strip_suffix(')')?.trim();
    if body.starts_with(['+', '-']) {
        return None;
    }
    if let Some(value) = parse_number(body) {
        return Some((-value, None));
    }
    parse_typed_number(body).map(|(value, id)| (-value, Some(id)))
}

fn parse_literal(input: &str) -> ParsedInput {
    if input.is_empty() {
        return ParsedInput::Clear;
    }
    if let Some(rest) = input.strip_prefix('=') {
        return ParsedInput::Formula(rest.to_string());
    }
    if let Some(rest) = input.strip_prefix('\'') {
        return ParsedInput::Text(rest.to_string());
    }
    if input.eq_ignore_ascii_case("true") {
        return ParsedInput::Bool(true);
    }
    if input.eq_ignore_ascii_case("false") {
        return ParsedInput::Bool(false);
    }
    if let Some(n) = parse_number(input) {
        return ParsedInput::Number(n);
    }
    ParsedInput::Text(input.to_string())
}

fn day_first(format: &FormatCode<'_>) -> bool {
    let code = match format {
        FormatCode::Builtin(id) => xlsx_model::numfmt::builtin_format_code(*id).unwrap_or(""),
        FormatCode::Custom(code) => code,
    };
    let mut quoted = false;
    let mut bracketed = false;
    let mut escaped = false;
    let mut first = None;
    let mut has_day = false;
    let mut has_month = false;
    for ch in code.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        if matches!(ch, '\\' | '_' | '*') && !quoted {
            escaped = true;
            continue;
        }
        match ch {
            '"' => quoted = !quoted,
            '[' if !quoted => bracketed = true,
            ']' if !quoted => bracketed = false,
            ';' if !quoted && !bracketed => break,
            _ if !quoted && !bracketed => match ch.to_ascii_lowercase() {
                'd' => {
                    has_day = true;
                    first.get_or_insert('d');
                }
                'm' => {
                    has_month = true;
                    first.get_or_insert('m');
                }
                _ => {}
            },
            _ => {}
        }
    }
    has_day && has_month && first == Some('d')
}

fn parse_grouped_number(input: &str) -> Option<f64> {
    let unsigned = input.strip_prefix(['+', '-']).unwrap_or(input);
    let (whole, fraction) = unsigned
        .split_once('.')
        .map_or((unsigned, None), |(a, b)| (a, Some(b)));
    if let Some(fraction) = fraction
        && (fraction.is_empty() || !fraction.bytes().all(|b| b.is_ascii_digit()))
    {
        return None;
    }
    let mut groups = whole.split(',');
    let first = groups.next()?;
    if first.is_empty() || first.len() > 3 || !first.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut count = 0;
    for group in groups {
        if group.len() != 3 || !group.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        count += 1;
    }
    if count == 0 {
        return None;
    }
    parse_number(&input.replace(',', ""))
}

fn parse_typed_number(input: &str) -> Option<(f64, u16)> {
    if let Some(body) = input.strip_suffix('%') {
        let value = parse_number(body).or_else(|| parse_grouped_number(body))? / 100.0;
        return Some((value, if body.contains('.') { 10 } else { 9 }));
    }
    let currency = input
        .strip_prefix('$')
        .map(str::to_owned)
        .or_else(|| input.strip_prefix("-$").map(|body| format!("-{body}")));
    if let Some(body) = currency {
        return parse_number(&body)
            .or_else(|| parse_grouped_number(&body))
            .map(|value| (value, 7));
    }
    parse_grouped_number(input).map(|value| (value, if input.contains('.') { 4 } else { 3 }))
}

fn parse_timestamp(
    input: &str,
    day_first: bool,
    date_system: DateSystem,
    now_serial: Option<f64>,
) -> Option<(f64, u16)> {
    let (date, time) = if let Some(colon) = input.find(':') {
        let start = input[..colon]
            .char_indices()
            .rev()
            .find(|(_, ch)| ch.is_whitespace())
            .map_or(0, |(index, ch)| index + ch.len_utf8());
        (input[..start].trim(), Some(&input[start..]))
    } else {
        (input, None)
    };
    let clock = match time {
        Some(time) => Some(parse_time(time)?),
        None => None,
    };
    if date.is_empty() {
        return clock;
    }
    let fields: Vec<_> = date
        .split(['/', '-', '.', ',', ' '])
        .filter(|s| !s.is_empty())
        .collect();
    if !(2..=3).contains(&fields.len())
        || !date
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '/' | '-' | '.' | ',' | ' '))
        || (fields.len() == 2 && !date.contains('/'))
        || (date.contains(',') && !date.bytes().any(|b| b.is_ascii_alphabetic()))
        || (date.contains('.') && (!day_first || fields.last().is_none_or(|year| year.len() != 4)))
    {
        return None;
    }
    let serial =
        xlsx_calc::functions::datetime::parse_entry_date(date, day_first, date_system, now_serial)?;
    Some(match clock {
        Some((fraction, _)) => (serial + fraction, 22),
        None => (serial, 14),
    })
}

fn parse_time(input: &str) -> Option<(f64, u16)> {
    let upper = input.to_ascii_uppercase();
    let am_pm = upper.ends_with("AM") || upper.ends_with("PM");
    let body = if am_pm {
        input[..input.len() - 2].trim_end()
    } else {
        input
    };
    let parts: Vec<_> = body.split(':').collect();
    if !(2..=3).contains(&parts.len())
        || parts
            .iter()
            .any(|p| p.is_empty() || !p.bytes().all(|b| b.is_ascii_digit()))
    {
        return None;
    }
    let hour: u32 = parts[0].parse().ok()?;
    let minute: u32 = parts[1].parse().ok()?;
    let second: u32 = parts.get(2).map_or(Some(0), |p| p.parse().ok())?;
    if minute >= 60
        || second >= 60
        || (!am_pm && hour >= 24)
        || (am_pm && !(1..=12).contains(&hour))
    {
        return None;
    }
    let hour = if am_pm {
        hour % 12 + if upper.ends_with("PM") { 12 } else { 0 }
    } else {
        hour
    };
    let value = f64::from(hour * 3600 + minute * 60 + second) / 86400.0;
    let id = match (am_pm, parts.len() == 3) {
        (false, false) => 20,
        (false, true) => 21,
        (true, false) => 18,
        (true, true) => 19,
    };
    Some((value, id))
}

/// parse a bare numeric literal; rejects the `inf`/`nan` tokens that
/// `f64::from_str` would accept, so those fall through to text.
fn parse_number(input: &str) -> Option<f64> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return None;
    }
    let n: f64 = trimmed.parse().ok()?;
    n.is_finite().then_some(n)
}

/// build the `CellState` for a raw editor string, evaluating formulas
/// immediately against `provider`; unparseable formulas degrade to raw text.
pub fn cell_state_for_input(input: &str, provider: &dyn CellProvider, sheet: SheetId) -> CellState {
    match parse_input(input) {
        ParsedInput::Clear => CellState::default(),
        ParsedInput::Number(value) => CellState {
            value: CellValue::Number { value },
            ..Default::default()
        },
        ParsedInput::Bool(value) => CellState {
            value: CellValue::Bool { value },
            ..Default::default()
        },
        ParsedInput::Text(value) => CellState {
            value: CellValue::Text { value },
            ..Default::default()
        },
        ParsedInput::Formula(text) => match parse_formula(&text) {
            Ok(expr) => {
                let ctx = EvalContext::new(provider, sheet);
                CellState {
                    value: evaluate(&expr, &ctx),
                    formula: Some(text),
                    ..Default::default()
                }
            }
            Err(_) => CellState {
                value: CellValue::Text {
                    value: input.to_string(),
                },
                ..Default::default()
            },
        },
    }
}

/// build the `CellState` without evaluating: formula text is stored with an
/// `Empty` value for the dependency graph's recalc to fill in afterwards.
pub fn cell_state_for_input_no_eval(input: &str) -> CellState {
    cell_state_from_parsed(input, parse_input(input))
}

pub fn cell_state_from_parsed(input: &str, parsed: ParsedInput) -> CellState {
    match parsed {
        ParsedInput::Clear => CellState::default(),
        ParsedInput::Number(value) => CellState {
            value: CellValue::Number { value },
            ..Default::default()
        },
        ParsedInput::Bool(value) => CellState {
            value: CellValue::Bool { value },
            ..Default::default()
        },
        ParsedInput::Text(value) => CellState {
            value: CellValue::Text { value },
            ..Default::default()
        },
        ParsedInput::Formula(text) => match parse_formula(&text) {
            Ok(_) => CellState {
                value: CellValue::Empty,
                formula: Some(text),
                ..Default::default()
            },
            Err(_) => CellState {
                value: CellValue::Text {
                    value: input.to_string(),
                },
                ..Default::default()
            },
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellRef, Sheet, Workbook};

    #[test]
    fn classifies_edge_cases() {
        assert_eq!(parse_input(""), ParsedInput::Clear);
        assert_eq!(parse_input("00"), ParsedInput::Number(0.0));
        assert_eq!(parse_input("1e3"), ParsedInput::Number(1000.0));
        assert_eq!(parse_input(" 42 "), ParsedInput::Number(42.0));
        assert_eq!(parse_input("="), ParsedInput::Formula(String::new()));
        assert_eq!(parse_input("'=danger"), ParsedInput::Text("=danger".into()));
        assert_eq!(parse_input("-"), ParsedInput::Text("-".into()));
        assert_eq!(parse_input("3.14.15"), ParsedInput::Text("3.14.15".into()));
    }

    #[test]
    fn booleans_are_case_insensitive() {
        assert_eq!(parse_input("true"), ParsedInput::Bool(true));
        assert_eq!(parse_input("TRUE"), ParsedInput::Bool(true));
        assert_eq!(parse_input("False"), ParsedInput::Bool(false));
    }

    #[test]
    fn rejects_non_finite_number_tokens() {
        assert_eq!(parse_input("inf"), ParsedInput::Text("inf".into()));
        assert_eq!(parse_input("nan"), ParsedInput::Text("nan".into()));
        assert_eq!(
            parse_input("infinity"),
            ParsedInput::Text("infinity".into())
        );
    }

    #[test]
    fn dates_times_and_numeric_entry_forms() {
        for (input, expected) in [
            ("1/15/2024", 45306.0),
            ("1/15/24", 45306.0),
            ("2024-01-15", 45306.0),
            ("15-Jan-2024", 45306.0),
            ("15-jAn-24", 45306.0),
            ("January 15, 2024", 45306.0),
            (" 1/15/2024 ", 45306.0),
            ("1/1/00", 36526.0),
            ("1/1/29", 47119.0),
            ("1/1/30", 10959.0),
            ("1/1/99", 36161.0),
            ("2/29/1900", 60.0),
            ("3/1/1900", 61.0),
            ("2/29/2024", 45351.0),
            ("12:00", 0.5),
            ("12:00:00", 0.5),
            ("12:00 AM", 0.0),
            ("12:00 PM", 0.5),
            ("1:30 pm", 13.5 / 24.0),
            ("1:30:15 AM", 5415.0 / 86400.0),
            ("1/15/2024 12:00 PM", 45306.5),
            ("Jan 15, 2024 12:00:00", 45306.5),
            ("15%", 0.15),
            ("12.5%", 0.125),
            ("$1,234.50", 1234.5),
            ("-$5", -5.0),
            ("1,234", 1234.0),
            ("1,234.5", 1234.5),
            (" 15% ", 0.15),
            ("(1234.5)", -1234.5),
            ("(1,234.5)", -1234.5),
            ("($5)", -5.0),
            ("(12.5%)", -0.125),
        ] {
            assert_eq!(parse_input(input), ParsedInput::Number(expected), "{input}");
        }
    }

    #[test]
    fn malformed_entries_remain_text() {
        for input in [
            "2/29/2023",
            "4/31/2024",
            "13/1/2024",
            "1/0/2024",
            "1904-02-30",
            "1/1/10000",
            "24:00",
            "12:60",
            "13:00 PM",
            "0:00 AM",
            "1:30:60",
            "1/15/2024 25:00",
            "1,23",
            "1234,567",
            "$nan",
            "inf%",
            "1/15",
            "hello\u{2003}12:00",
            "(-5)",
            "()",
            "(5",
        ] {
            assert_eq!(
                parse_input(input),
                ParsedInput::Text(input.into()),
                "{input}"
            );
        }
        assert_eq!(parse_input("'15%"), ParsedInput::Text("15%".into()));
        assert_eq!(
            parse_input_with_context("(1234.5)", FormatCode::Builtin(0), DateSystem::V1900, None),
            (ParsedInput::Number(-1234.5), None)
        );
    }

    #[test]
    fn entry_uses_effective_date_format_epoch_and_clock() {
        let parse =
            |input, format, system, now| parse_input_with_context(input, format, system, now);
        assert_eq!(
            parse(
                "15/1/2024",
                FormatCode::Custom("dd/mm/yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45306.0), None)
        );
        assert_eq!(
            parse(
                "15.01.2024",
                FormatCode::Custom("dd.mm.yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45306.0), None)
        );
        assert_eq!(
            parse(
                "1.2.3",
                FormatCode::Custom("dd.mm.yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Text("1.2.3".into()), None)
        );
        assert_eq!(
            parse(
                "2024-01-15",
                FormatCode::Custom("dd/mm/yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45306.0), None)
        );
        assert_eq!(
            parse(
                "1/2/2024",
                FormatCode::Custom("dd/mm/yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45323.0), None)
        );
        assert_eq!(
            parse(
                "1/2/2024",
                FormatCode::Custom("\"day\" m/d/yyyy"),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45293.0), None)
        );
        assert_eq!(
            parse(
                "1/15/2024",
                FormatCode::Builtin(14),
                DateSystem::V1900,
                None
            ),
            (ParsedInput::Number(45306.0), None)
        );
        assert_eq!(
            parse("1/15/2024", FormatCode::Builtin(0), DateSystem::V1904, None),
            (ParsedInput::Number(43844.0), Some(14))
        );
        assert_eq!(
            parse("1/1/1904", FormatCode::Builtin(0), DateSystem::V1904, None),
            (ParsedInput::Number(0.0), Some(14))
        );
        assert_eq!(
            parse("2/29/1900", FormatCode::Builtin(0), DateSystem::V1904, None),
            (ParsedInput::Text("2/29/1900".into()), None)
        );
        assert_eq!(
            parse(
                "1/15",
                FormatCode::Builtin(0),
                DateSystem::V1900,
                Some(45306.0)
            ),
            (ParsedInput::Number(45306.0), Some(14))
        );
        assert_eq!(
            parse(
                "15/1",
                FormatCode::Custom("d/m/yyyy"),
                DateSystem::V1904,
                Some(43844.0)
            ),
            (ParsedInput::Number(43844.0), None)
        );
        assert_eq!(
            parse(
                "2/29",
                FormatCode::Builtin(0),
                DateSystem::V1900,
                Some(44927.0)
            ),
            (ParsedInput::Text("2/29".into()), None)
        );
    }

    #[test]
    fn general_entries_get_matching_builtin_formats() {
        for (input, id) in [
            ("1/15/2024", 14),
            ("1/15/2024 12:00", 22),
            ("12:00", 20),
            ("12:00:00", 21),
            ("12:00 AM", 18),
            ("12:00:00 PM", 19),
            ("15%", 9),
            ("12.5%", 10),
            ("$1,234.50", 7),
            ("-$5", 7),
            ("1,234", 3),
            ("1,234.5", 4),
            ("(1,234.5)", 4),
            ("($5)", 7),
        ] {
            let (_, format) =
                parse_input_with_context(input, FormatCode::Builtin(0), DateSystem::V1900, None);
            assert_eq!(format, Some(id), "{input}");
            assert_eq!(
                parse_input_with_context(input, FormatCode::Builtin(2), DateSystem::V1900, None).1,
                None,
                "{input}"
            );
        }
    }

    #[test]
    fn typed_numbers_in_general_cells() {
        fn general(input: &str) -> (ParsedInput, Option<u16>) {
            parse_input_with_context(input, FormatCode::Builtin(0), DateSystem::V1900, None)
        }
        let number = |value: f64, id: Option<u16>| (ParsedInput::Number(value), id);
        let text = |input: &str| (ParsedInput::Text(input.into()), None);
        for (input, expected) in [
            ("1,234.5", number(1234.5, Some(4))),
            ("  1,234.5  ", number(1234.5, Some(4))),
            ("-1,234", number(-1234.0, Some(3))),
            ("+1,234,567", number(1_234_567.0, Some(3))),
            ("12.5%", number(0.125, Some(10))),
            (" 12.5% ", number(0.125, Some(10))),
            ("1,234%", number(12.34, Some(9))),
            ("-12.5%", number(-0.125, Some(10))),
            ("(1234.5)", number(-1234.5, None)),
            ("( 1234.5 )", number(-1234.5, None)),
            ("(1,234)", number(-1234.0, Some(3))),
            ("  42  ", number(42.0, None)),
            ("00123", number(123.0, None)),
            ("1,2,3", text("1,2,3")),
            ("1,2,2024", text("1,2,2024")),
            ("12,31,2024", text("12,31,2024")),
            ("1,23", text("1,23")),
            ("12,3456", text("12,3456")),
            ("1,234,56", text("1,234,56")),
            ("1234,567.5", text("1234,567.5")),
            (",123", text(",123")),
            ("1,234.", text("1,234.")),
            ("1,,234", text("1,,234")),
            ("1,234 5", text("1,234 5")),
            ("%12", text("%12")),
            ("12%%", text("12%%")),
            ("(1,234", text("(1,234")),
            ("'1,234.5", text("1,234.5")),
            ("'12.5%", text("12.5%")),
            ("'(1234.5)", text("(1234.5)")),
            ("Jan 2, 2003", number(37623.0, Some(14))),
            ("January 15, 2024", number(45306.0, Some(14))),
            ("15 Jan, 2024", number(45306.0, Some(14))),
            ("1-2-3", number(37623.0, Some(14))),
            ("1-2-2024", number(45293.0, Some(14))),
            ("1/2/2024", number(45293.0, Some(14))),
        ] {
            assert_eq!(general(input), expected, "{input}");
        }
    }

    #[test]
    fn text_formats_keep_literal_input() {
        for format in [FormatCode::Builtin(49), FormatCode::Custom("@")] {
            for input in [
                "1/15/2024",
                "15%",
                "$1,234.50",
                "12:00",
                "=1+1",
                "TRUE",
                "'15%",
            ] {
                assert_eq!(
                    parse_input_with_context(input, format.clone(), DateSystem::V1900, None),
                    (
                        ParsedInput::Text(input.strip_prefix('\'').unwrap_or(input).into()),
                        None
                    ),
                );
            }
        }
    }

    fn wb_with(cells: &[(&str, f64)]) -> Workbook {
        let mut sheet = Sheet::new("Sheet1");
        for (a1, v) in cells {
            sheet.set_cell(
                CellRef::parse_a1(a1).unwrap(),
                Cell {
                    value: CellValue::Number { value: *v },
                    ..Cell::default()
                },
            );
        }
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        wb
    }

    #[test]
    fn formula_evaluates_against_provider_on_entry() {
        let wb = wb_with(&[("A1", 2.0), ("A2", 3.0)]);
        let state = cell_state_for_input("=SUM(A1:A2)", &wb, SheetId(0));
        assert_eq!(state.value, CellValue::Number { value: 5.0 });
        assert_eq!(state.formula.as_deref(), Some("SUM(A1:A2)"));
    }

    #[test]
    fn currency_formula_entry_matches_recalc() {
        let mut wb = wb_with(&[]);
        wb.styles.num_fmts.push((164, "£0.00".into()));
        let state = cell_state_for_input("=\"£12\"+1", &wb, SheetId(0));
        assert_eq!(state.value, CellValue::Number { value: 13.0 });
        let at = CellRef::new(0, 0);
        wb.sheets[0].set_cell(
            at,
            xlsx_model::Cell {
                formula: state.formula,
                ..Default::default()
            },
        );
        xlsx_calc::rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(state.value, wb.sheets[0].cell(at).unwrap().value);
    }

    #[test]
    fn unparseable_formula_degrades_to_raw_text() {
        let wb = wb_with(&[]);
        let state = cell_state_for_input("=SUM(", &wb, SheetId(0));
        assert_eq!(
            state.value,
            CellValue::Text {
                value: "=SUM(".into()
            }
        );
        assert!(state.formula.is_none());
    }

    #[test]
    fn empty_input_clears() {
        let wb = wb_with(&[]);
        assert_eq!(
            cell_state_for_input("", &wb, SheetId(0)),
            CellState::default()
        );
    }

    #[test]
    fn no_eval_stores_formula_without_value() {
        let state = cell_state_for_input_no_eval("=SUM(A1:A2)");
        assert_eq!(state.value, CellValue::Empty);
        assert_eq!(state.formula.as_deref(), Some("SUM(A1:A2)"));
    }

    #[test]
    fn no_eval_coerces_literals_and_degrades_bad_formula() {
        assert_eq!(
            cell_state_for_input_no_eval("42").value,
            CellValue::Number { value: 42.0 }
        );
        let bad = cell_state_for_input_no_eval("=SUM(");
        assert_eq!(
            bad.value,
            CellValue::Text {
                value: "=SUM(".into()
            }
        );
        assert!(bad.formula.is_none());
    }
}
