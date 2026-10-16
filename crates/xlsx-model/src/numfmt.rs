//! number-format code interpreter (ecma-376 §18.8.30–31): value + format code
//! -> display string and optional bracket color. unsupported degrades to general.

use std::cell::RefCell;
use std::collections::HashMap;
use std::sync::Arc;

use serde::{Deserialize, Serialize};

use crate::date::DateSystem;
use crate::value::CellValue;

/// the display string plus the `#rrggbb` color a bracket prefix (`[Red]`) requested.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct FormattedValue {
    pub text: String,
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout: Vec<FormatPart>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum FormatPartKind {
    Padding,
    Fill,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct FormatPart {
    pub offset: usize,
    pub kind: FormatPartKind,
    pub character: char,
}

impl FormattedValue {
    fn push_str(&mut self, text: &str) {
        self.text.push_str(text);
    }

    fn push(&mut self, character: char) {
        self.text.push(character);
    }

    fn padding(&mut self, character: char) {
        self.layout.push(FormatPart {
            offset: self.text.len(),
            kind: FormatPartKind::Padding,
            character,
        });
        self.text.push(' ');
    }

    fn fill(&mut self, character: char) {
        if !self
            .layout
            .iter()
            .any(|part| part.kind == FormatPartKind::Fill)
        {
            self.layout.push(FormatPart {
                offset: self.text.len(),
                kind: FormatPartKind::Fill,
                character,
            });
        }
    }

    fn digits(&mut self, text: &str) {
        if !text.contains(' ') {
            self.push_str(text);
            return;
        }
        for character in text.chars() {
            if character == ' ' {
                self.padding('0');
            } else {
                self.push(character);
            }
        }
    }

    fn append(&mut self, mut other: Self) {
        for part in &mut other.layout {
            part.offset += self.text.len();
        }
        self.push_str(&other.text);
        self.layout.extend(other.layout);
    }
}

impl From<String> for FormattedValue {
    fn from(text: String) -> Self {
        Self {
            text,
            ..Self::default()
        }
    }
}

/// format a cell value with a number-format code. bools, errors, and empties
/// ignore the code; text uses the 4th section or an `@` placeholder.
pub fn format_value(value: &CellValue, code: &str, date_system: DateSystem) -> FormattedValue {
    match value {
        CellValue::Bool { value } => FormattedValue {
            text: if *value { "TRUE" } else { "FALSE" }.to_string(),
            color: None,
            layout: Vec::new(),
        },
        CellValue::Error { value } => FormattedValue {
            text: value.as_str().to_string(),
            color: None,
            layout: Vec::new(),
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
        5 => r##""$"#,##0_);("$"#,##0)"##,
        6 => r##""$"#,##0_);[Red]("$"#,##0)"##,
        7 => r##""$"#,##0.00_);("$"#,##0.00)"##,
        8 => r##""$"#,##0.00_);[Red]("$"#,##0.00)"##,
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        41 => r##"_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)"##,
        42 => r##"_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)"##,
        43 => r##"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"##,
        44 => r##"_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)"##,
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
    Width(char),
    Fill(char),
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

#[derive(Clone, Debug, Default, PartialEq)]
struct Section {
    toks: Vec<Tok>,
    color: Option<String>,
    condition: Option<Condition>,
    fraction: Option<FractionPattern>,
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

/// memoized parses per generation; capped so adversarial code variety can't grow the cache.
const FORMAT_CACHE_CAP: usize = 128;

/// codes longer than this parse fresh every call and are never retained; excel
/// itself rejects format codes beyond ~255 chars.
const MAX_CACHED_CODE_LEN: usize = 256;

/// positive, negative, zero, text: the only sections any render path selects.
const USED_SECTIONS: usize = 4;

/// two generations, so making room costs one swap instead of a scan for a
/// victim: inserts fill `hot`, a full `hot` ages into `cold`, and the next
/// age-out drops it. holds at most `2 * FORMAT_CACHE_CAP` parses.
#[derive(Default)]
struct FormatCache {
    hot: HashMap<String, Arc<[Section]>>,
    cold: HashMap<String, Arc<[Section]>>,
}

impl FormatCache {
    fn get(&mut self, code: &str) -> Option<Arc<[Section]>> {
        if let Some(hit) = self.hot.get(code) {
            return Some(hit.clone());
        }
        let promoted = self.cold.get(code)?.clone();
        self.insert(code, promoted.clone());
        Some(promoted)
    }

    fn insert(&mut self, code: &str, parsed: Arc<[Section]>) {
        if self.hot.len() >= FORMAT_CACHE_CAP {
            self.cold = std::mem::take(&mut self.hot);
        }
        self.hot.insert(code.to_string(), parsed);
    }
}

thread_local! {
    /// per thread, so concurrent renders never serialize on a shared lock.
    static FORMAT_CACHE: RefCell<FormatCache> = RefCell::new(FormatCache::default());
}

/// memoized `split_sections` + `tokenize`, keyed by the exact code string.
/// falls back to parsing when the thread's cache is already torn down.
fn parsed_sections(code: &str) -> Arc<[Section]> {
    FORMAT_CACHE
        .try_with(|cache| parsed_sections_in(&mut cache.borrow_mut(), code))
        .unwrap_or_else(|_| tokenize_all(code))
}

fn parsed_sections_in(cache: &mut FormatCache, code: &str) -> Arc<[Section]> {
    if code.len() > MAX_CACHED_CODE_LEN {
        return tokenize_all(code);
    }
    if let Some(hit) = cache.get(code) {
        return hit;
    }
    let parsed = tokenize_all(code);
    cache.insert(code, parsed.clone());
    parsed
}

/// `select` never looks past the third section and text never past the fourth,
/// so deeper ones are dropped rather than parsed and retained.
fn tokenize_all(code: &str) -> Arc<[Section]> {
    split_sections(code)
        .iter()
        .take(USED_SECTIONS)
        .map(|s| tokenize(s))
        .collect()
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
                    section.toks.push(Tok::Width(chars[i]));
                    i += 1;
                }
            }
            '*' => {
                i += 1;
                if i < n {
                    if !section.has(|t| matches!(t, Tok::Fill(_))) {
                        section.toks.push(Tok::Fill(chars[i]));
                    }
                    i += 1;
                }
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
    section.fraction = fraction_pattern(&section.toks);
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
    } else {
        (Cmp::Eq, s.strip_prefix('=')?)
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
            layout: Vec::new(),
        };
    }
    let parsed = parsed_sections(code);
    let (idx, auto_minus) = select(&parsed, n);
    let sec = &parsed[idx];
    let mut formatted = if sec.has_date() {
        render_date_section(sec, n, ds)
    } else {
        let mut body = render_number_section(sec, n.abs());
        if auto_minus {
            body.text.insert(0, '-');
            for part in &mut body.layout {
                part.offset += 1;
            }
        }
        body
    };
    formatted.color = sec.color.clone();
    formatted
}

/// whether formatting `value` with `code` degrades to a general render
/// because the code asks for something the interpreter stubs or a fraction
/// does not fit.
pub fn format_is_approximate(value: &CellValue, code: &str) -> bool {
    let CellValue::Number { value } = value else {
        return false;
    };
    let trimmed = code.trim();
    if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("general") {
        return false;
    }
    let parsed = parsed_sections(code);
    let section = &parsed[select(&parsed, *value).0];
    !section.has_date()
        && (degrades_to_general(section) || fraction_falls_back(section, value.abs()))
}

fn fraction_falls_back(sec: &Section, mag: f64) -> bool {
    let Some(pattern) = sec.fraction else {
        return false;
    };
    !sec.has(|t| matches!(t, Tok::General | Tok::Exp(_))) && fraction_parts(mag, pattern).is_none()
}

fn degrades_to_general(sec: &Section) -> bool {
    !sec.has(|t| matches!(t, Tok::General))
        && sec.has(|t| matches!(t, Tok::Slash))
        && sec.fraction.is_none()
}

fn render_number_section(sec: &Section, mag: f64) -> FormattedValue {
    if sec.has(|t| matches!(t, Tok::General)) {
        return render_general_section(sec, mag);
    }
    if degrades_to_general(sec) {
        return format_general(mag).into();
    }
    if sec.has(|t| matches!(t, Tok::Exp(_))) {
        return render_scientific(sec, mag);
    }
    if let Some(pattern) = sec.fraction {
        return render_fraction(sec, mag, pattern);
    }
    render_decimal_tokens(&sec.toks, mag, true)
}

fn render_decimal_tokens(toks: &[Tok], mag: f64, apply_scaling: bool) -> FormattedValue {
    let total_ph = toks.iter().filter(|t| matches!(t, Tok::Digit(_))).count();
    let has_at = toks.iter().any(|t| matches!(t, Tok::At));
    if total_ph == 0 && !has_at {
        return emit_literals_only(toks);
    }

    let dot_pos = toks.iter().position(|t| matches!(t, Tok::Dot));
    let (int_toks, frac_toks): (&[Tok], &[Tok]) = match dot_pos {
        Some(p) => (&toks[..p], &toks[p + 1..]),
        None => (toks, &[]),
    };

    let (grouping, scale) = scan_commas(int_toks);
    let percent = toks.iter().filter(|t| matches!(t, Tok::Percent)).count() as i32;

    let mut v = mag;
    if apply_scaling {
        for _ in 0..percent {
            v *= 100.0;
        }
        for _ in 0..scale {
            v /= 1000.0;
        }
    }

    let int_kinds: Vec<char> = digit_kinds(int_toks);
    let frac_kinds: Vec<char> = digit_kinds(frac_toks);
    let (int_digits, frac_digits) = split_digits(v, frac_kinds.len());
    let int_number = build_int_number(&int_kinds, &int_digits, grouping);
    let frac_str = build_frac(&frac_kinds, &frac_digits);

    let mut out = FormattedValue::default();
    let mut int_cursor = 0;
    let mut int_placeholder = 0;
    let mut frac_placeholder = 0;
    let mut passed_dot = false;
    for t in toks {
        match t {
            Tok::Digit(_) => {
                if !passed_dot {
                    int_placeholder += 1;
                    let remaining = int_kinds[int_placeholder..]
                        .iter()
                        .enumerate()
                        .filter(|(index, kind)| {
                            int_digits.len() > int_kinds.len() - int_placeholder - index - 1
                                || matches!(**kind, '0' | '?')
                        })
                        .count();
                    let end = int_number.len().saturating_sub(remaining);
                    out.digits(&int_number[int_cursor..end]);
                    int_cursor = end;
                } else {
                    let j = frac_placeholder;
                    let significant = (j..frac_kinds.len())
                        .any(|k| frac_digits.as_bytes()[k] != b'0' || frac_kinds[k] == '0');
                    if significant {
                        out.push(frac_digits.as_bytes()[j] as char);
                    } else if frac_kinds[j] == '?' {
                        out.padding('0');
                    }
                    frac_placeholder += 1;
                }
            }
            Tok::Dot => {
                passed_dot = true;
                if int_kinds.is_empty() {
                    out.digits(&int_number);
                }
                if !frac_str.is_empty() {
                    out.push('.');
                }
            }
            Tok::Comma => {}
            Tok::Percent => out.push('%'),
            Tok::At => out.push_str(&format_general(mag)),
            Tok::Literal(s) => out.push_str(s),
            Tok::Width(c) => out.padding(*c),
            Tok::Fill(c) => out.fill(*c),
            _ => {}
        }
    }
    out
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct FractionPattern {
    numerator: usize,
    slash: usize,
    denominator_end: usize,
    fixed: Option<u64>,
    limit: u64,
    mixed: bool,
}

fn fraction_pattern(toks: &[Tok]) -> Option<FractionPattern> {
    let slash = toks.iter().position(|t| matches!(t, Tok::Slash))?;
    let mut numerator = slash;
    while numerator > 0 && matches!(toks[numerator - 1], Tok::Digit(_)) {
        numerator -= 1;
    }
    if numerator == slash {
        return None;
    }
    let mut end = slash + 1;
    let mut literal = String::new();
    while let Some(tok) = toks.get(end) {
        match tok {
            Tok::Literal(s) if s.chars().all(|c| c.is_ascii_digit()) => literal.push_str(s),
            Tok::Digit('0') => literal.push('0'),
            _ => break,
        }
        end += 1;
    }
    let fixed = if !literal.is_empty() {
        Some(literal.parse::<u64>().ok().filter(|value| *value > 0)?)
    } else {
        while matches!(toks.get(end), Some(Tok::Digit(_))) {
            end += 1;
        }
        None
    };
    let digits = end - slash - 1;
    if digits == 0 || digits > 9 {
        return None;
    }
    Some(FractionPattern {
        numerator,
        slash,
        denominator_end: end,
        fixed,
        limit: 10u64.pow(digits as u32) - 1,
        mixed: !digit_kinds(&toks[..numerator]).is_empty(),
    })
}

/// closest fraction within `limit`, or `None` when a convergent's numerator
/// does not fit a u64.
fn best_fraction(value: f64, limit: u64) -> Option<(u64, u64)> {
    let (mut p0, mut q0, mut p1, mut q1) = (0u64, 1u64, 1u64, 0u64);
    let mut remainder = value;
    loop {
        let whole = remainder.floor() as u64;
        let Some(q2) = whole.checked_mul(q1).and_then(|q| q0.checked_add(q)) else {
            break;
        };
        if q2 > limit {
            break;
        }
        let p2 = whole.checked_mul(p1)?.checked_add(p0)?;
        (p0, q0, p1, q1) = (p1, q1, p2, q2);
        let fraction = remainder - whole as f64;
        if fraction <= f64::EPSILON {
            return Some((p1, q1));
        }
        remainder = 1.0 / fraction;
    }
    if q1 == 0 {
        return None;
    }
    let count = (limit - q0) / q1;
    let (p2, q2) = (count.checked_mul(p1)?.checked_add(p0)?, q0 + count * q1);
    if (p1 as f64 / q1 as f64 - value).abs() <= (p2 as f64 / q2 as f64 - value).abs() {
        Some((p1, q1))
    } else {
        Some((p2, q2))
    }
}

/// (whole, numerator, denominator) for `mag`, or `None` when one does not
/// fit a u64 and the section shows general text instead.
fn fraction_parts(mag: f64, pattern: FractionPattern) -> Option<(u64, u64, u64)> {
    if !mag.is_finite() || mag >= u64::MAX as f64 {
        return None;
    }
    let mut whole = if pattern.mixed { mag.floor() as u64 } else { 0 };
    let value = if pattern.mixed {
        mag - whole as f64
    } else {
        mag
    };
    let (mut numerator, denominator) = match pattern.fixed {
        Some(denominator) => {
            let scaled = (value * denominator as f64).round();
            if scaled >= u64::MAX as f64 {
                return None;
            }
            (scaled as u64, denominator)
        }
        None => best_fraction(value, pattern.limit)?,
    };
    if pattern.mixed && numerator >= denominator {
        whole += numerator / denominator;
        numerator %= denominator;
    }
    Some((whole, numerator, denominator))
}

fn render_fraction(sec: &Section, mag: f64, pattern: FractionPattern) -> FormattedValue {
    let Some((whole, numerator, denominator)) = fraction_parts(mag, pattern) else {
        return format_general(mag).into();
    };
    let numerator_kinds = digit_kinds(&sec.toks[pattern.numerator..pattern.slash]);
    let denominator_kinds = digit_kinds(&sec.toks[pattern.slash + 1..pattern.denominator_end]);
    let hidden = pattern.mixed && numerator == 0 && !numerator_kinds.contains(&'0');
    let mut out = render_decimal_tokens(&sec.toks[..pattern.numerator], whole as f64, false);
    if hidden {
        for _ in &numerator_kinds {
            out.padding('0');
        }
        out.padding('/');
        let denominator_width = if pattern.fixed.is_some() {
            denominator.to_string().len()
        } else {
            denominator_kinds.len()
        };
        for _ in 0..denominator_width {
            out.padding('0');
        }
    } else {
        out.digits(&build_int_number(
            &numerator_kinds,
            &numerator.to_string(),
            false,
        ));
        out.push('/');
        let digits = denominator.to_string();
        out.push_str(&digits);
        if pattern.fixed.is_none() {
            for kind in denominator_kinds.iter().skip(digits.len()) {
                if *kind == '?' {
                    out.padding('0');
                } else if *kind == '0' {
                    out.push('0');
                }
            }
        }
    }
    for tok in &sec.toks[pattern.denominator_end..] {
        match tok {
            Tok::Literal(s) => out.push_str(s),
            Tok::Width(c) => out.padding(*c),
            Tok::Fill(c) => out.fill(*c),
            Tok::Percent => out.push('%'),
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

fn emit_literals_only(toks: &[Tok]) -> FormattedValue {
    let mut out = FormattedValue::default();
    for t in toks {
        match t {
            Tok::Literal(s) => out.push_str(s),
            Tok::Width(c) => out.padding(*c),
            Tok::Fill(c) => out.fill(*c),
            _ => {}
        }
    }
    out
}

fn render_general_section(sec: &Section, mag: f64) -> FormattedValue {
    let mut out = FormattedValue::default();
    for t in &sec.toks {
        match t {
            Tok::General => out.push_str(&format_general(mag)),
            Tok::Literal(s) => out.push_str(s),
            Tok::Width(c) => out.padding(*c),
            Tok::Fill(c) => out.fill(*c),
            _ => {}
        }
    }
    out
}

/// `0.00E+00` / engineering `##0.0E+0`: the count of integer placeholders sets
/// the exponent step, so the mantissa's integer part fills 1..=p digits.
fn render_scientific(sec: &Section, mag: f64) -> FormattedValue {
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

    let (mantissa, exp_val) = if v == 0.0 {
        (0.0, 0i32)
    } else {
        let e = v.log10().floor() as i32;
        let mut exp = e.div_euclid(p) * p;
        let mut m = v / 10f64.powi(exp);
        // rounding can overflow the mantissa (9.99 -> 10): bump exponent, recompute
        let (id, _) = split_digits(m, qf);
        if id.len() as i32 > p {
            exp += p;
            m = v / 10f64.powi(exp);
        }
        (m, exp)
    };

    let sign = if exp_val < 0 {
        "-"
    } else if plus {
        "+"
    } else {
        ""
    };
    let mut out = render_decimal_tokens(mant_toks, mantissa, false);
    out.push_str(&format!("E{sign}"));
    out.append(render_decimal_tokens(
        exp_toks,
        exp_val.unsigned_abs() as f64,
        false,
    ));
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
        scientific_precision(a, 5)
    } else {
        let decimals = (10 - e).clamp(0, 15) as usize;
        let mut s = round_fixed(a, decimals);
        trim_trailing_zeros(&mut s);
        s
    };
    if neg { format!("-{body}") } else { body }
}

/// `a` in scientific notation with `decimals` mantissa places, rounded half
/// away from zero from its first fifteen significant digits as Excel shows it.
fn scientific_precision(a: f64, decimals: usize) -> String {
    let exact = format!("{a:.14e}");
    let (mantissa, exponent) = exact.split_once('e').unwrap_or((exact.as_str(), "0"));
    let mut exponent: i32 = exponent.parse().unwrap_or(0);
    let mut digits: Vec<u8> = mantissa.bytes().filter(u8::is_ascii_digit).collect();
    let round_up = digits.get(decimals + 1).is_some_and(|digit| *digit >= b'5');
    digits.truncate(decimals + 1);
    if round_up {
        match digits.iter().rposition(|digit| *digit != b'9') {
            Some(last) => {
                digits[last] += 1;
                digits[last + 1..].fill(b'0');
            }
            None => {
                digits.fill(b'0');
                digits[0] = b'1';
                exponent += 1;
            }
        }
    }
    let mut m = String::with_capacity(digits.len() + 1);
    m.push(char::from(digits[0]));
    if digits.len() > 1 {
        m.push('.');
        m.extend(digits[1..].iter().map(|digit| char::from(*digit)));
    }
    trim_trailing_zeros(&mut m);
    let sign = if exponent < 0 { '-' } else { '+' };
    format!("{m}E{sign}{:02}", exponent.unsigned_abs())
}

pub fn is_general(value: &CellValue, code: &str) -> bool {
    let CellValue::Number { value } = value else {
        return false;
    };
    let code = code.trim();
    if code.is_empty() || code.eq_ignore_ascii_case("general") {
        return true;
    }
    let parsed = parsed_sections(code);
    parsed[select(&parsed, *value).0].toks == [Tok::General]
}

pub fn fit_general(value: f64, width: f32, mut measure: impl FnMut(&str) -> f32) -> Option<String> {
    if !value.is_finite() || value == 0.0 {
        return None;
    }
    let exponent = value.abs().log10().floor() as i32;
    if (-4..11).contains(&exponent) {
        for decimals in (0..=(10 - exponent).clamp(0, 15) as usize).rev() {
            let mut candidate = round_fixed(value, decimals);
            trim_trailing_zeros(&mut candidate);
            if candidate.parse::<f64>().ok() != Some(0.0) && measure(&candidate) <= width {
                return Some(candidate);
            }
        }
    }
    for decimals in (0..=5).rev() {
        let body = scientific_precision(value.abs(), decimals);
        let candidate = if value < 0.0 {
            format!("-{body}")
        } else {
            body
        };
        if measure(&candidate) <= width {
            return Some(candidate);
        }
    }
    None
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
    let parsed = parsed_sections(code);
    let sec = if parsed.len() >= 4 {
        Some(&parsed[3])
    } else {
        parsed.iter().find(|s| s.has(|t| matches!(t, Tok::At)))
    };
    match sec {
        Some(s) => {
            let mut out = FormattedValue::default();
            for t in &s.toks {
                match t {
                    Tok::At => out.push_str(text),
                    Tok::Literal(l) => out.push_str(l),
                    Tok::Width(c) => out.padding(*c),
                    Tok::Fill(c) => out.fill(*c),
                    _ => {}
                }
            }
            out.color = s.color.clone();
            out
        }
        None => FormattedValue {
            text: text.to_string(),
            color: None,
            layout: Vec::new(),
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
fn render_date_section(sec: &Section, serial: f64, ds: DateSystem) -> FormattedValue {
    if serial < 0.0 {
        return "#######".to_string().into();
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
        None => return "#######".to_string().into(),
    };
    let weekday = weekday_sun0(year, month, day);
    let ampm = sec.has(|t| matches!(t, Tok::AmPm(_, _)));
    let minute_flags = minute_flags(&sec.toks);

    let mut out = FormattedValue::default();
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
            Tok::Width(c) => out.padding(*c),
            Tok::Fill(c) => out.fill(*c),
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
