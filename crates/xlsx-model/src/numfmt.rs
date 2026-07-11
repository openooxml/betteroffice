//! number-format code interpreter (ecma-376 §18.8.30–31): value + format code
//! -> display string and optional bracket color. unsupported degrades to general.

use serde::{Deserialize, Serialize};

use crate::date::DateSystem;
use crate::value::CellValue;

/// the display string plus the `#rrggbb` color a bracket prefix (`[Red]`) requested.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct FormattedValue {
    pub text: String,
    pub color: Option<String>,
}

/// format a cell value with a number-format code. bools, errors, and empties
/// ignore the code; text uses the 4th section or an `@` placeholder.
pub fn format_value(value: &CellValue, code: &str, date_system: DateSystem) -> FormattedValue {
    match value {
        CellValue::Bool { value } => FormattedValue {
            text: if *value { "TRUE" } else { "FALSE" }.to_string(),
            color: None,
        },
        CellValue::Error { value } => FormattedValue {
            text: value.as_str().to_string(),
            color: None,
        },
        CellValue::Empty => FormattedValue::default(),
        CellValue::Text { value } => format_text_value(value, code),
        CellValue::Number { value } => format_number_value(*value, code, date_system),
    }
}

/// the implied builtin number-format table (§18.8.30); locale-dependent ids
/// return None so callers fall back to general.
pub fn builtin_format_code(id: u16) -> Option<&'static str> {
    let code = match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    };
    Some(code)
}

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    Digit(char),
    Dot,
    Comma,
    Percent,
    Exp(bool),
    Slash,
    At,
    General,
    Literal(String),
    Width,
    Fill,
    Year(u8),
    Mon(u8), // month or minute, resolved by adjacency
    Day(u8),
    Hour(u8),
    Sec(u8),
    Elapsed(char, u8),
    AmPm(String, String),
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Cmp {
    Lt,
    Le,
    Gt,
    Ge,
    Eq,
    Ne,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct Condition {
    op: Cmp,
    value: f64,
}

impl Condition {
    fn matches(&self, n: f64) -> bool {
        match self.op {
            Cmp::Lt => n < self.value,
            Cmp::Le => n <= self.value,
            Cmp::Gt => n > self.value,
            Cmp::Ge => n >= self.value,
            Cmp::Eq => n == self.value,
            Cmp::Ne => n != self.value,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct Section {
    toks: Vec<Tok>,
    color: Option<String>,
    condition: Option<Condition>,
}

impl Section {
    fn has_date(&self) -> bool {
        self.toks.iter().any(|t| {
            matches!(
                t,
                Tok::Year(_)
                    | Tok::Mon(_)
                    | Tok::Day(_)
                    | Tok::Hour(_)
                    | Tok::Sec(_)
                    | Tok::Elapsed(_, _)
                    | Tok::AmPm(_, _)
            )
        })
    }

    fn has(&self, pred: impl Fn(&Tok) -> bool) -> bool {
        self.toks.iter().any(pred)
    }
}

/// split a code into sections on top-level `;` (quotes, brackets, and escapes guarded).
fn split_sections(code: &str) -> Vec<String> {
    let chars: Vec<char> = code.chars().collect();
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' => {
                cur.push(c);
                i += 1;
                while i < chars.len() {
                    cur.push(chars[i]);
                    let done = chars[i] == '"';
                    i += 1;
                    if done {
                        break;
                    }
                }
            }
            '[' => {
                cur.push(c);
                i += 1;
                while i < chars.len() {
                    cur.push(chars[i]);
                    let done = chars[i] == ']';
                    i += 1;
                    if done {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                cur.push(c);
                i += 1;
                if i < chars.len() {
                    cur.push(chars[i]);
                    i += 1;
                }
            }
            ';' => {
                out.push(std::mem::take(&mut cur));
                i += 1;
            }
            _ => {
                cur.push(c);
                i += 1;
            }
        }
    }
    out.push(cur);
    out
}

fn tokenize(input: &str) -> Section {
    let chars: Vec<char> = input.chars().collect();
    let n = chars.len();
    let mut section = Section::default();
    let mut i = 0;
    while i < n {
        let c = chars[i];
        match c {
            '"' => {
                i += 1;
                let mut s = String::new();
                while i < n && chars[i] != '"' {
                    s.push(chars[i]);
                    i += 1;
                }
                if i < n {
                    i += 1;
                }
                section.toks.push(Tok::Literal(s));
            }
            '\\' => {
                i += 1;
                if i < n {
                    section.toks.push(Tok::Literal(chars[i].to_string()));
                    i += 1;
                }
            }
            '_' => {
                i += 1;
                if i < n {
                    i += 1;
                }
                section.toks.push(Tok::Width);
            }
            '*' => {
                i += 1;
                if i < n {
                    i += 1;
                }
                section.toks.push(Tok::Fill);
            }
            '[' => {
                i += 1;
                let mut inner = String::new();
                while i < n && chars[i] != ']' {
                    inner.push(chars[i]);
                    i += 1;
                }
                if i < n {
                    i += 1;
                }
                parse_bracket(&inner, &mut section);
            }
            '0' | '#' | '?' => {
                section.toks.push(Tok::Digit(c));
                i += 1;
            }
            '.' => {
                section.toks.push(Tok::Dot);
                i += 1;
            }
            ',' => {
                section.toks.push(Tok::Comma);
                i += 1;
            }
            '%' => {
                section.toks.push(Tok::Percent);
                i += 1;
            }
            '/' => {
                section.toks.push(Tok::Slash);
                i += 1;
            }
            '@' => {
                section.toks.push(Tok::At);
                i += 1;
            }
            'E' | 'e' if i + 1 < n && (chars[i + 1] == '+' || chars[i + 1] == '-') => {
                section.toks.push(Tok::Exp(chars[i + 1] == '+'));
                i += 2;
            }
            'y' | 'Y' => {
                let cnt = run_len(&chars, i);
                section.toks.push(Tok::Year(cnt as u8));
                i += cnt;
            }
            'm' | 'M' => {
                let cnt = run_len(&chars, i);
                section.toks.push(Tok::Mon(cnt as u8));
                i += cnt;
            }
            'd' | 'D' => {
                let cnt = run_len(&chars, i);
                section.toks.push(Tok::Day(cnt as u8));
                i += cnt;
            }
            'h' | 'H' => {
                let cnt = run_len(&chars, i);
                section.toks.push(Tok::Hour(cnt as u8));
                i += cnt;
            }
            's' | 'S' => {
                let cnt = run_len(&chars, i);
                section.toks.push(Tok::Sec(cnt as u8));
                i += cnt;
            }
            'a' | 'A' => match match_ampm(&chars, i) {
                Some((am, pm, len)) => {
                    section.toks.push(Tok::AmPm(am, pm));
                    i += len;
                }
                None => {
                    section.toks.push(Tok::Literal(c.to_string()));
                    i += 1;
                }
            },
            'g' | 'G' if match_general(&chars, i) => {
                section.toks.push(Tok::General);
                i += 7;
            }
            _ => {
                section.toks.push(Tok::Literal(c.to_string()));
                i += 1;
            }
        }
    }
    section
}

fn run_len(chars: &[char], i: usize) -> usize {
    let c = chars[i].to_ascii_lowercase();
    let mut j = i;
    while j < chars.len() && chars[j].to_ascii_lowercase() == c {
        j += 1;
    }
    j - i
}

fn match_ampm(chars: &[char], i: usize) -> Option<(String, String, usize)> {
    let rest: String = chars[i..].iter().collect();
    let low = rest.to_ascii_lowercase();
    if low.starts_with("am/pm") {
        let am: String = chars[i..i + 2].iter().collect();
        let pm: String = chars[i + 3..i + 5].iter().collect();
        Some((am, pm, 5))
    } else if low.starts_with("a/p") {
        let am: String = chars[i..i + 1].iter().collect();
        let pm: String = chars[i + 2..i + 3].iter().collect();
        Some((am, pm, 3))
    } else {
        None
    }
}

fn match_general(chars: &[char], i: usize) -> bool {
    let rest: String = chars[i..].iter().collect();
    rest.to_ascii_lowercase().starts_with("general")
}

/// interpret a `[...]` prefix: color, condition, elapsed-time field, or
/// `[$sym-locale]` currency (strip locale, keep symbol). unknown -> drop.
fn parse_bracket(inner: &str, section: &mut Section) {
    let trimmed = inner.trim();
    if let Some(hex) = color_hex(trimmed) {
        section.color = Some(hex);
        return;
    }
    if let Some(cond) = parse_condition(trimmed) {
        section.condition = Some(cond);
        return;
    }
    if let Some(tok) = parse_elapsed(trimmed) {
        section.toks.push(tok);
        return;
    }
    if let Some(rest) = trimmed.strip_prefix('$') {
        let sym = rest.split_once('-').map_or(rest, |(s, _)| s);
        if !sym.is_empty() {
            section.toks.push(Tok::Literal(sym.to_string()));
        }
    }
}

fn color_hex(s: &str) -> Option<String> {
    let up = s.to_ascii_uppercase();
    let hex = match up.as_str() {
        "BLACK" => "#000000",
        "WHITE" => "#FFFFFF",
        "RED" => "#FF0000",
        "GREEN" => "#00FF00",
        "BLUE" => "#0000FF",
        "YELLOW" => "#FFFF00",
        "MAGENTA" => "#FF00FF",
        "CYAN" => "#00FFFF",
        _ => {
            let rest = up.strip_prefix("COLOR")?;
            let idx: u32 = rest.trim().parse().ok()?;
            return indexed_color(idx);
        }
    };
    Some(hex.to_string())
}

/// the classic `[Color N]` palette, indices 1–8 only; higher indices are None.
fn indexed_color(idx: u32) -> Option<String> {
    let hex = match idx {
        1 => "#000000",
        2 => "#FFFFFF",
        3 => "#FF0000",
        4 => "#00FF00",
        5 => "#0000FF",
        6 => "#FFFF00",
        7 => "#FF00FF",
        8 => "#00FFFF",
        _ => return None,
    };
    Some(hex.to_string())
}

fn parse_condition(s: &str) -> Option<Condition> {
    let (op, rest) = if let Some(r) = s.strip_prefix(">=") {
        (Cmp::Ge, r)
    } else if let Some(r) = s.strip_prefix("<=") {
        (Cmp::Le, r)
    } else if let Some(r) = s.strip_prefix("<>") {
        (Cmp::Ne, r)
    } else if let Some(r) = s.strip_prefix('>') {
        (Cmp::Gt, r)
    } else if let Some(r) = s.strip_prefix('<') {
        (Cmp::Lt, r)
    } else if let Some(r) = s.strip_prefix('=') {
        (Cmp::Eq, r)
    } else {
        return None;
    };
    let value: f64 = rest.trim().parse().ok()?;
    Some(Condition { op, value })
}

fn parse_elapsed(s: &str) -> Option<Tok> {
    let c0 = s.chars().next()?.to_ascii_lowercase();
    if !matches!(c0, 'h' | 'm' | 's') {
        return None;
    }
    if !s.chars().all(|c| c.to_ascii_lowercase() == c0) {
        return None;
    }
    Some(Tok::Elapsed(c0, s.chars().count() as u8))
}

/// pick the section index for a number and whether to prepend a `-`:
/// excel's sign fallback, overridden by explicit conditions tried in order.
fn select(parsed: &[Section], n: f64) -> (usize, bool) {
    let cap = parsed.len().min(3);
    let has_cond = parsed[..cap].iter().any(|s| s.condition.is_some());
    if has_cond {
        for (i, s) in parsed.iter().enumerate() {
            if i >= 3 {
                break;
            }
            match &s.condition {
                Some(c) => {
                    if c.matches(n) {
                        return (i, false);
                    }
                }
                None => return (i, n < 0.0),
            }
        }
        return (0, n < 0.0);
    }
    let count = parsed.len();
    if n > 0.0 {
        (0, false)
    } else if n == 0.0 {
        if count >= 3 { (2, false) } else { (0, false) }
    } else if count >= 2 {
        (1, false)
    } else {
        (0, true)
    }
}

fn format_number_value(n: f64, code: &str, ds: DateSystem) -> FormattedValue {
    let trimmed = code.trim();
    if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("general") {
        return FormattedValue {
            text: format_general(n),
            color: None,
        };
    }
    let parsed: Vec<Section> = split_sections(code).iter().map(|s| tokenize(s)).collect();
    let (idx, auto_minus) = select(&parsed, n);
    let sec = &parsed[idx];
    let color = sec.color.clone();
    let text = if sec.has_date() {
        render_date_section(sec, n, ds)
    } else {
        let body = render_number_section(sec, n.abs());
        if auto_minus { format!("-{body}") } else { body }
    };
    FormattedValue { text, color }
}

fn render_number_section(sec: &Section, mag: f64) -> String {
    if sec.has(|t| matches!(t, Tok::General)) {
        return render_general_section(sec, mag);
    }
    // fractions are stubbed: degrade to a general render.
    if sec.has(|t| matches!(t, Tok::Slash)) {
        return format_general(mag);
    }
    if sec.has(|t| matches!(t, Tok::Exp(_))) {
        return render_scientific(sec, mag);
    }
    let total_ph = sec
        .toks
        .iter()
        .filter(|t| matches!(t, Tok::Digit(_)))
        .count();
    let has_at = sec.has(|t| matches!(t, Tok::At));
    if total_ph == 0 && !has_at {
        return emit_literals_only(sec);
    }

    let dot_pos = sec.toks.iter().position(|t| matches!(t, Tok::Dot));
    let (int_toks, frac_toks): (&[Tok], &[Tok]) = match dot_pos {
        Some(p) => (&sec.toks[..p], &sec.toks[p + 1..]),
        None => (&sec.toks[..], &[]),
    };

    let (grouping, scale) = scan_commas(int_toks);
    let percent = sec
        .toks
        .iter()
        .filter(|t| matches!(t, Tok::Percent))
        .count() as i32;

    let mut v = mag;
    for _ in 0..percent {
        v *= 100.0;
    }
    for _ in 0..scale {
        v /= 1000.0;
    }

    let int_kinds: Vec<char> = digit_kinds(int_toks);
    let frac_kinds: Vec<char> = digit_kinds(frac_toks);
    let (int_digits, frac_digits) = split_digits(v, frac_kinds.len());
    let int_number = build_int_number(&int_kinds, &int_digits, grouping);
    let frac_str = build_frac(&frac_kinds, &frac_digits);

    let mut out = String::new();
    let mut int_emitted = false;
    let mut passed_dot = false;
    for t in &sec.toks {
        match t {
            Tok::Digit(_) => {
                if !passed_dot && !int_emitted {
                    out.push_str(&int_number);
                    int_emitted = true;
                }
            }
            Tok::Dot => {
                passed_dot = true;
                if !int_emitted {
                    out.push_str(&int_number);
                    int_emitted = true;
                }
                if !frac_str.is_empty() {
                    out.push('.');
                    out.push_str(&frac_str);
                }
            }
            Tok::Comma => {}
            Tok::Percent => out.push('%'),
            Tok::At => out.push_str(&format_general(mag)),
            Tok::Literal(s) => out.push_str(s),
            Tok::Width => out.push(' '),
            Tok::Fill => {}
            _ => {}
        }
    }
    out
}

/// classify integer-side commas: flanked by digits turns on thousands
/// grouping; trailing ones scale the value down by 1000 each.
fn scan_commas(int_toks: &[Tok]) -> (bool, u32) {
    let digit_positions: Vec<usize> = int_toks
        .iter()
        .enumerate()
        .filter(|(_, t)| matches!(t, Tok::Digit(_)))
        .map(|(i, _)| i)
        .collect();
    let mut grouping = false;
    let mut scale = 0u32;
    for (idx, t) in int_toks.iter().enumerate() {
        if !matches!(t, Tok::Comma) {
            continue;
        }
        let before = digit_positions.iter().any(|&p| p < idx);
        let after = digit_positions.iter().any(|&p| p > idx);
        if before && after {
            grouping = true;
        } else if before {
            scale += 1;
        }
    }
    (grouping, scale)
}

fn digit_kinds(toks: &[Tok]) -> Vec<char> {
    toks.iter()
        .filter_map(|t| match t {
            Tok::Digit(c) => Some(*c),
            _ => None,
        })
        .collect()
}

/// split a non-negative magnitude into (integer digits, fraction digits) after
/// rounding half-away-from-zero (excel's display rule) to `frac_len` places.
fn split_digits(mag: f64, frac_len: usize) -> (String, String) {
    let factor = 10f64.powi(frac_len as i32);
    let scaled = (mag * factor).round();
    let mut s = format!("{scaled:.0}");
    while s.len() < frac_len + 1 {
        s.insert(0, '0');
    }
    let cut = s.len() - frac_len;
    let int_part = &s[..cut];
    let frac_part = s[cut..].to_string();
    let int_trimmed = int_part.trim_start_matches('0');
    (int_trimmed.to_string(), frac_part)
}

fn build_int_number(kinds: &[char], int_digits: &str, grouping: bool) -> String {
    let dchars: Vec<char> = int_digits.chars().collect();
    let n = dchars.len();
    let p = kinds.len();
    if p == 0 {
        return group_if(int_digits, grouping);
    }
    let mut buf: Vec<char> = Vec::new();
    for (j, &kind) in kinds.iter().enumerate() {
        let rev = p - 1 - j;
        if j == 0 {
            let take = n.saturating_sub(rev);
            if take == 0 {
                if let Some(ch) = pad_char(kind) {
                    buf.push(ch);
                }
            } else {
                buf.extend_from_slice(&dchars[0..take]);
            }
        } else {
            let pos = n as isize - 1 - rev as isize;
            if pos >= 0 {
                buf.push(dchars[pos as usize]);
            } else if let Some(ch) = pad_char(kind) {
                buf.push(ch);
            }
        }
    }
    let s: String = buf.into_iter().collect();
    group_if(&s, grouping)
}

fn pad_char(kind: char) -> Option<char> {
    match kind {
        '0' => Some('0'),
        '?' => Some(' '),
        _ => None,
    }
}

fn group_if(s: &str, grouping: bool) -> String {
    if !grouping {
        return s.to_string();
    }
    let lead = s.chars().take_while(|c| *c == ' ').count();
    let (spaces, rest) = s.split_at(lead);
    format!("{spaces}{}", group_thousands(rest))
}

fn group_thousands(digits: &str) -> String {
    let n = digits.len();
    let mut out = String::with_capacity(n + n / 3);
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (n - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(c);
    }
    out
}

/// place fraction digits into placeholders, suppressing trailing insignificant
/// ones: `#` drops, `?` pads a space, `0` keeps the zero.
fn build_frac(kinds: &[char], frac_digits: &str) -> String {
    let fchars: Vec<char> = frac_digits.chars().collect();
    let q = kinds.len();
    let mut out = String::new();
    for j in 0..q {
        let significant = (j..q).any(|k| fchars[k] != '0' || kinds[k] == '0');
        if significant {
            out.push(fchars[j]);
        } else if kinds[j] == '?' {
            out.push(' ');
        }
    }
    out
}

fn emit_literals_only(sec: &Section) -> String {
    let mut out = String::new();
    for t in &sec.toks {
        match t {
            Tok::Literal(s) => out.push_str(s),
            Tok::Width => out.push(' '),
            _ => {}
        }
    }
    out
}

fn render_general_section(sec: &Section, mag: f64) -> String {
    let mut out = String::new();
    for t in &sec.toks {
        match t {
            Tok::General => out.push_str(&format_general(mag)),
            Tok::Literal(s) => out.push_str(s),
            Tok::Width => out.push(' '),
            _ => {}
        }
    }
    out
}

/// `0.00E+00` / engineering `##0.0E+0`: the count of integer placeholders sets
/// the exponent step, so the mantissa's integer part fills 1..=p digits.
fn render_scientific(sec: &Section, mag: f64) -> String {
    let exp_idx = sec
        .toks
        .iter()
        .position(|t| matches!(t, Tok::Exp(_)))
        .unwrap();
    let plus = matches!(sec.toks[exp_idx], Tok::Exp(true));
    let mant_toks = &sec.toks[..exp_idx];
    let exp_toks = &sec.toks[exp_idx + 1..];

    let mant_dot = mant_toks.iter().position(|t| matches!(t, Tok::Dot));
    let (mi, mf): (&[Tok], &[Tok]) = match mant_dot {
        Some(p) => (&mant_toks[..p], &mant_toks[p + 1..]),
        None => (mant_toks, &[]),
    };
    let int_kinds = digit_kinds(mi);
    let frac_kinds = digit_kinds(mf);
    let p = int_kinds.len().max(1) as i32;
    let qf = frac_kinds.len();

    let percent = sec
        .toks
        .iter()
        .filter(|t| matches!(t, Tok::Percent))
        .count() as i32;
    let mut v = mag;
    for _ in 0..percent {
        v *= 100.0;
    }

    let (mant_out, exp_val) = if v == 0.0 {
        (build_mantissa(&int_kinds, &frac_kinds, 0.0), 0i32)
    } else {
        let e = v.log10().floor() as i32;
        let mut exp = e.div_euclid(p) * p;
        let mut m = v / 10f64.powi(exp);
        let mut mant = build_mantissa(&int_kinds, &frac_kinds, m);
        // rounding can overflow the mantissa (9.99 -> 10): bump exponent, recompute
        let (id, _) = split_digits(m, qf);
        if id.len() as i32 > p {
            exp += p;
            m = v / 10f64.powi(exp);
            mant = build_mantissa(&int_kinds, &frac_kinds, m);
        }
        (mant, exp)
    };

    let exp_width = exp_toks
        .iter()
        .filter(|t| matches!(t, Tok::Digit('0')))
        .count()
        .max(1);
    let sign = if exp_val < 0 {
        "-"
    } else if plus {
        "+"
    } else {
        ""
    };
    let exp_digits = format!("{:0width$}", exp_val.unsigned_abs(), width = exp_width);

    let prefix = leading_literals(mant_toks);
    let suffix = trailing_literals(exp_toks);
    format!("{prefix}{mant_out}E{sign}{exp_digits}{suffix}")
}

fn build_mantissa(int_kinds: &[char], frac_kinds: &[char], m: f64) -> String {
    let (int_digits, frac_digits) = split_digits(m, frac_kinds.len());
    let int_number = build_int_number(int_kinds, &int_digits, false);
    let frac_str = build_frac(frac_kinds, &frac_digits);
    if frac_str.is_empty() {
        int_number
    } else {
        format!("{int_number}.{frac_str}")
    }
}

fn leading_literals(toks: &[Tok]) -> String {
    let mut out = String::new();
    for t in toks {
        match t {
            Tok::Literal(s) => out.push_str(s),
            Tok::Digit(_) | Tok::Dot => break,
            _ => {}
        }
    }
    out
}

fn trailing_literals(toks: &[Tok]) -> String {
    let mut out = String::new();
    for t in toks {
        if let Tok::Literal(s) = t {
            out.push_str(s);
        }
    }
    out
}

/// excel's "general" render: signed, ~11 significant digits, switching to
/// scientific outside roughly 1e-4 .. 1e11.
fn format_general(n: f64) -> String {
    if !n.is_finite() {
        return if n.is_nan() {
            "NaN".to_string()
        } else if n > 0.0 {
            "INF".to_string()
        } else {
            "-INF".to_string()
        };
    }
    if n == 0.0 {
        return "0".to_string();
    }
    let neg = n < 0.0;
    let a = n.abs();
    let e = a.log10().floor() as i32;
    let body = if !(-4..11).contains(&e) {
        general_scientific(a, e)
    } else {
        let decimals = (10 - e).clamp(0, 15) as usize;
        let mut s = round_fixed(a, decimals);
        trim_trailing_zeros(&mut s);
        s
    };
    if neg { format!("-{body}") } else { body }
}

fn general_scientific(a: f64, e: i32) -> String {
    let mant = a / 10f64.powi(e);
    let mut m = round_fixed(mant, 5);
    trim_trailing_zeros(&mut m);
    let sign = if e < 0 { '-' } else { '+' };
    format!("{m}E{sign}{:02}", e.unsigned_abs())
}

fn round_fixed(a: f64, decimals: usize) -> String {
    let factor = 10f64.powi(decimals as i32);
    let r = (a * factor).round() / factor;
    format!("{r:.decimals$}")
}

fn trim_trailing_zeros(s: &mut String) {
    if s.contains('.') {
        while s.ends_with('0') {
            s.pop();
        }
        if s.ends_with('.') {
            s.pop();
        }
    }
}

fn format_text_value(text: &str, code: &str) -> FormattedValue {
    let parsed: Vec<Section> = split_sections(code).iter().map(|s| tokenize(s)).collect();
    let sec = if parsed.len() >= 4 {
        Some(&parsed[3])
    } else {
        parsed.iter().find(|s| s.has(|t| matches!(t, Tok::At)))
    };
    match sec {
        Some(s) => {
            let mut out = String::new();
            for t in &s.toks {
                match t {
                    Tok::At => out.push_str(text),
                    Tok::Literal(l) => out.push_str(l),
                    Tok::Width => out.push(' '),
                    _ => {}
                }
            }
            FormattedValue {
                text: out,
                color: s.color.clone(),
            }
        }
        None => FormattedValue {
            text: text.to_string(),
            color: None,
        },
    }
}

const MONTH_FULL: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

const DAY_FULL: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

/// render a serial through a date/time section; negatives show excel's `#######`.
fn render_date_section(sec: &Section, serial: f64, ds: DateSystem) -> String {
    if serial < 0.0 {
        return "#######".to_string();
    }
    let secfrac = fractional_second_digits(&sec.toks);
    let sub = 10i128.pow(secfrac);
    let total_subunits = (serial * 86400.0 * sub as f64).round() as i128;
    let day_span = 86_400i128 * sub;
    let day_serial = total_subunits.div_euclid(day_span) as i64;
    let within = total_subunits.rem_euclid(day_span);
    let whole_seconds = (within / sub) as i64;
    let frac_units = (within % sub) as i64;
    let hour24 = whole_seconds / 3600;
    let minute = (whole_seconds % 3600) / 60;
    let second = whole_seconds % 60;
    let elapsed_seconds = (total_subunits / sub) as i64;

    let (year, month, day) = match serial_to_civil(day_serial, ds) {
        Some(v) => v,
        None => return "#######".to_string(),
    };
    let weekday = weekday_sun0(year, month, day);
    let ampm = sec.has(|t| matches!(t, Tok::AmPm(_, _)));
    let minute_flags = minute_flags(&sec.toks);

    let mut out = String::new();
    let mut i = 0;
    while i < sec.toks.len() {
        match &sec.toks[i] {
            Tok::Year(c) => out.push_str(&fmt_year(year, *c)),
            Tok::Mon(c) => {
                if minute_flags[i] {
                    out.push_str(&fmt_two(minute, *c));
                } else {
                    out.push_str(&fmt_month(month, *c));
                }
            }
            Tok::Day(c) => out.push_str(&fmt_day(day, weekday, *c)),
            Tok::Hour(c) => {
                let h = if ampm {
                    ((hour24 + 11) % 12) + 1
                } else {
                    hour24
                };
                out.push_str(&fmt_two(h, *c));
            }
            Tok::Sec(c) => out.push_str(&fmt_two(second, *c)),
            Tok::Elapsed(ch, c) => {
                let val = match ch {
                    'h' => elapsed_seconds / 3600,
                    'm' => elapsed_seconds / 60,
                    _ => elapsed_seconds,
                };
                out.push_str(&format!("{val:0width$}", width = *c as usize));
            }
            Tok::AmPm(am, pm) => out.push_str(if hour24 < 12 { am } else { pm }),
            Tok::Dot => {
                let mut cnt = 0;
                let mut j = i + 1;
                while j < sec.toks.len() && matches!(sec.toks[j], Tok::Digit(_)) {
                    cnt += 1;
                    j += 1;
                }
                if cnt > 0 {
                    let padded = format!("{frac_units:0width$}", width = secfrac as usize);
                    out.push('.');
                    out.push_str(&padded[..cnt.min(padded.len())]);
                    i = j;
                    continue;
                }
                out.push('.');
            }
            // in a date section, slash and comma are literal separators, not number ops
            Tok::Slash => out.push('/'),
            Tok::Comma => out.push(','),
            Tok::Literal(s) => out.push_str(s),
            Tok::Width => out.push(' '),
            _ => {}
        }
        i += 1;
    }
    out
}

/// fractional-second precision: digit count directly after the first `.`.
fn fractional_second_digits(toks: &[Tok]) -> u32 {
    for (k, t) in toks.iter().enumerate() {
        if !matches!(t, Tok::Dot) {
            continue;
        }
        let mut cnt = 0u32;
        let mut j = k + 1;
        while j < toks.len() && matches!(toks[j], Tok::Digit(_)) {
            cnt += 1;
            j += 1;
        }
        if cnt > 0 {
            return cnt;
        }
    }
    0
}

/// mark which `Mon` tokens are minutes: an `m` run next to an hour (before)
/// or seconds (after) field is minutes, else a month.
fn minute_flags(toks: &[Tok]) -> Vec<bool> {
    let mut cats: Vec<(usize, char)> = Vec::new();
    for (k, t) in toks.iter().enumerate() {
        match t {
            Tok::Hour(_) => cats.push((k, 'h')),
            Tok::Mon(_) => cats.push((k, 'm')),
            Tok::Sec(_) => cats.push((k, 's')),
            Tok::Elapsed(c, _) => cats.push((k, *c)),
            _ => {}
        }
    }
    let mut flags = vec![false; toks.len()];
    for (pos, (k, c)) in cats.iter().enumerate() {
        if *c != 'm' {
            continue;
        }
        let prev = pos.checked_sub(1).map(|p| cats[p].1);
        let next = cats.get(pos + 1).map(|x| x.1);
        if prev == Some('h') || next == Some('s') {
            flags[*k] = true;
        }
    }
    flags
}

fn fmt_year(y: i64, c: u8) -> String {
    if c <= 2 {
        format!("{:02}", y.rem_euclid(100))
    } else {
        format!("{y:04}")
    }
}

fn fmt_two(v: i64, c: u8) -> String {
    if c <= 1 {
        format!("{v}")
    } else {
        format!("{v:02}")
    }
}

fn fmt_month(m: i64, c: u8) -> String {
    let idx = (m.clamp(1, 12) - 1) as usize;
    match c {
        1 => format!("{m}"),
        2 => format!("{m:02}"),
        3 => MONTH_FULL[idx][..3].to_string(),
        4 => MONTH_FULL[idx].to_string(),
        _ => MONTH_FULL[idx][..1].to_string(),
    }
}

fn fmt_day(d: i64, weekday: usize, c: u8) -> String {
    match c {
        1 => format!("{d}"),
        2 => format!("{d:02}"),
        3 => DAY_FULL[weekday][..3].to_string(),
        _ => DAY_FULL[weekday].to_string(),
    }
}

fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// (year, month, day) for an integer serial. 1900 system: serial 60 is the
/// phantom 1900-02-29, serial 0 is excel's 1900-01-00.
fn serial_to_civil(serial: i64, ds: DateSystem) -> Option<(i64, i64, i64)> {
    match ds {
        DateSystem::V1900 => {
            if serial < 0 {
                None
            } else if serial == 0 {
                Some((1900, 1, 0))
            } else if serial == 60 {
                Some((1900, 2, 29))
            } else {
                let adjusted = if serial > 60 { serial - 1 } else { serial };
                Some(civil_from_days(adjusted - 25_568))
            }
        }
        DateSystem::V1904 => {
            if serial < 0 {
                None
            } else {
                Some(civil_from_days(serial - 24_107))
            }
        }
    }
}

/// day of week with Sunday = 0 (proleptic gregorian; 1970-01-01 was Thursday).
fn weekday_sun0(y: i64, m: i64, d: i64) -> usize {
    (days_from_civil(y, m, d) + 4).rem_euclid(7) as usize
}

#[cfg(test)]
mod tests;
