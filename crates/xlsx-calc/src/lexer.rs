//! formula lexer: turns source (without the leading `=`) into position-tagged
//! tokens, resolving reference/range/sheet disambiguation.

use std::fmt;

pub(crate) use xlsx_model::sheet_qualifier::can_start_span as is_unquoted_sheet;
use xlsx_model::sheet_qualifier::{is_path_qualifier, parse_sheet_token};
use xlsx_model::{CellRange, CellRef, ErrorValue};

use crate::reference::MAX_TABLE_SPEC_ITEMS;
use crate::{ColumnRange, RowRange, TableBand, TableSpec};

pub const MAX_TOKENS: usize = 10_000;
pub const MAX_FORMULA_BYTES: usize = 32_768;
pub const MAX_FORMULA_CHARS: usize = 8192;
const TABLE_ESCAPE_CHARS: &[char] = &['[', ']', '#', '@', '\''];
/// how many significant digits of a number literal excel keeps; the rest
/// read as zeros.
const LITERAL_DIGITS: usize = 15;
/// the smallest literal excel refuses; 9.99999999999999E+307 is the largest
/// it accepts.
const LITERAL_LIMIT: f64 = 1e308;
/// the smallest nonzero literal excel is known to accept: 1E-309 is taken
/// and 1E-310 refused; where between the two the boundary lies is not known.
const LITERAL_SMALLEST: f64 = 1e-309;

/// a literal's value inside excel's range: past either end there is none,
/// and one below the smallest normal double reads as 0, as Excel rewrites it.
fn literal_in_range(value: f64) -> Option<f64> {
    if value >= LITERAL_LIMIT || (value != 0.0 && value < LITERAL_SMALLEST) {
        return None;
    }
    Some(if value < f64::MIN_POSITIVE {
        0.0
    } else {
        value
    })
}

/// a number literal as excel reads it: digits past the fifteenth significant
/// one become zeros before the text is converted.
fn literal_value(text: &str) -> Option<f64> {
    let split = text.find(['e', 'E']).unwrap_or(text.len());
    let (mantissa, exponent) = text.split_at(split);
    let mut kept = String::with_capacity(text.len());
    let mut significant = 0;
    for c in mantissa.chars() {
        if c.is_ascii_digit() && (significant > 0 || c != '0') {
            significant += 1;
            kept.push(if significant > LITERAL_DIGITS { '0' } else { c });
        } else {
            kept.push(c);
        }
    }
    kept.push_str(exponent);
    kept.parse().ok()
}

/// a positioned lexer/parser error, the single error type of `parse_formula`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseError {
    /// byte offset into the source where the problem starts.
    pub pos: usize,
    pub message: String,
}

impl ParseError {
    pub fn new(pos: usize, message: impl Into<String>) -> Self {
        Self {
            pos,
            message: message.into(),
        }
    }
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "formula error at {}: {}", self.pos, self.message)
    }
}

impl std::error::Error for ParseError {}

/// error literals recognized in formulas, longest-unambiguous set from the spec.
const ERROR_LITERALS: &[(&str, ErrorValue)] = &[
    ("#DIV/0!", ErrorValue::Div0),
    ("#N/A", ErrorValue::NA),
    ("#NAME?", ErrorValue::Name),
    ("#NULL!", ErrorValue::Null),
    ("#NUM!", ErrorValue::Num),
    ("#REF!", ErrorValue::Ref),
    ("#VALUE!", ErrorValue::Value),
    ("#SPILL!", ErrorValue::Spill),
    ("#CALC!", ErrorValue::Calc),
];

#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokKind,
    pub start: usize,
    pub end: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokKind {
    Num(f64),
    Str(String),
    Bool(bool),
    ErrLit(ErrorValue),
    /// a bare word that is not a bool/ref — a function name or defined name.
    Ident(String),
    Name {
        scope: Option<String>,
        name: String,
    },
    Ref {
        sheet: Option<String>,
        cell: CellRef,
    },
    Range {
        sheet: Option<String>,
        range: CellRange,
    },
    ColumnRange {
        sheet: Option<String>,
        range: ColumnRange,
    },
    RowRange {
        sheet: Option<String>,
        range: RowRange,
    },
    TableRef {
        sheet: Option<String>,
        table: String,
        spec: TableSpec,
    },
    ThreeDRef {
        first: String,
        last: String,
        target: Box<TokKind>,
    },
    ExternalRef(Box<TokKind>),
    Intersection,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Percent,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    LParen,
    RParen,
    Comma,
    /// the range operator between two references, as in `A1:INDEX(..)`.
    /// a literal `A1:B2` is one `Range` token and never reaches here.
    Colon,
    LBrace,
    RBrace,
    Semicolon,
}

/// tokenize a formula source string.
pub fn lex(input: &str) -> Result<Vec<Token>, ParseError> {
    lex_with_range_names(input, is_formula_parameter)
}

/// Tokenize with caller-supplied names that can start a range join.
#[doc(hidden)]
pub fn lex_with_range_names(
    input: &str,
    is_range_name: impl Fn(&str) -> bool,
) -> Result<Vec<Token>, ParseError> {
    Lexer {
        input,
        pos: 0,
        table_refs: None,
        is_range_name: &is_range_name,
        is_sheet_span: None,
        evaluation_error: None,
    }
    .run()
}

pub(crate) type SheetSpan<'a> = dyn Fn(&str, &str) -> bool + 'a;

pub(crate) fn lex_for_evaluation(
    input: &str,
    is_sheet_span: Option<&SheetSpan<'_>>,
) -> Result<Vec<Token>, ParseError> {
    let mut error = None;
    let tokens = Lexer {
        input,
        pos: 0,
        table_refs: None,
        is_range_name: &is_formula_parameter,
        is_sheet_span,
        evaluation_error: Some(&mut error),
    }
    .run()?;
    error.map_or(Ok(tokens), Err)
}

fn is_formula_parameter(name: &str) -> bool {
    name.split_once('.')
        .is_some_and(|(prefix, _)| prefix.eq_ignore_ascii_case("_xlpm"))
}

/// Tokenize formula text for rewrites.
pub fn lex_for_rewrite(input: &str) -> Result<Vec<Token>, ParseError> {
    lex(input)
}

#[doc(hidden)]
pub fn ensure_unambiguous_spaced_calls(input: &str, tokens: &[Token]) -> Result<(), ParseError> {
    check_spaced_calls(input, tokens, false)
}

#[doc(hidden)]
pub fn ensure_unambiguous_spaced_references(
    input: &str,
    tokens: &[Token],
) -> Result<(), ParseError> {
    check_spaced_calls(input, tokens, true)
}

fn check_spaced_calls(
    input: &str,
    tokens: &[Token],
    allow_numeric_call: bool,
) -> Result<(), ParseError> {
    for (index, paren) in tokens.iter().enumerate() {
        if paren.kind != TokKind::LParen || index == 0 {
            continue;
        }
        let previous = &tokens[index - 1];
        let reference = if previous.kind == TokKind::Intersection && index >= 2 {
            &tokens[index - 2]
        } else {
            previous
        };
        if reference.end == paren.start {
            continue;
        }
        let kind = match &reference.kind {
            TokKind::ThreeDRef { target, .. } => target.as_ref(),
            kind => kind,
        };
        if !matches!(
            kind,
            TokKind::Ref { .. } | TokKind::Range { .. } | TokKind::ColumnRange { .. }
        ) {
            continue;
        }
        let endpoint = input[reference.start..reference.end]
            .rsplit(['!', ':'])
            .next()
            .unwrap_or_default()
            .trim_start();
        let builtin = crate::functions::resolve(endpoint);
        if builtin.is_none() && !crate::array::is_array_builtin(endpoint) {
            continue;
        }
        let group = &tokens[index + 1..];
        if matches!(kind, TokKind::ColumnRange { .. })
            && looks_like_range_group(group, builtin == Some(crate::functions::Func::N))
        {
            continue;
        }
        if allow_numeric_call
            && matches!(reference.kind, TokKind::Ref { sheet: None, .. })
            && matches!(
                group,
                [
                    Token {
                        kind: TokKind::Num(_),
                        ..
                    },
                    Token {
                        kind: TokKind::RParen,
                        ..
                    },
                    ..
                ]
            )
        {
            continue;
        }
        return Err(ParseError::new(
            reference.start,
            "ambiguous spaced function call",
        ));
    }
    Ok(())
}

fn looks_like_range_group(tokens: &[Token], allow_union: bool) -> bool {
    if let [first, close, ..] = tokens
        && close.kind == TokKind::RParen
    {
        return matches!(
            first.kind,
            TokKind::Ref { .. }
                | TokKind::Range { .. }
                | TokKind::ColumnRange { .. }
                | TokKind::RowRange { .. }
        );
    }
    if !allow_union {
        return false;
    }
    let (pairs, _) = tokens.as_chunks::<2>();
    for [reference, separator] in pairs {
        if !matches!(
            reference.kind,
            TokKind::Range { .. } | TokKind::ColumnRange { .. } | TokKind::RowRange { .. }
        ) {
            return false;
        }
        match separator.kind {
            TokKind::Comma => {}
            TokKind::RParen => return true,
            _ => return false,
        }
    }
    false
}

#[doc(hidden)]
pub fn rewrite_table_column(
    input: &str,
    table: &str,
    sheet: &str,
    local: bool,
    old: &str,
    new: &str,
) -> Result<Option<String>, ParseError> {
    if !may_reference_table_column(input, table, local, old) {
        return Ok(None);
    }
    let table = table.to_lowercase();
    let sheet = sheet.to_lowercase();
    let old = old.to_lowercase();
    rewrite_table_columns(input, |scope, name, column| {
        let same_table = if name.is_empty() {
            local
        } else {
            name.to_lowercase() == table && scope.is_none_or(|scope| scope.to_lowercase() == sheet)
        };
        (same_table && column.to_lowercase() == old).then_some(new)
    })
}

#[doc(hidden)]
pub fn may_reference_table_column(input: &str, table: &str, local: bool, column: &str) -> bool {
    if !input.contains('[') {
        return false;
    }
    let table = format!("{}[", table.to_lowercase());
    let column = column.to_lowercase();
    let unescaped_column = column.replace('\'', "");
    let mut has_selector = false;
    let mut has_column = false;
    let mut candidate = |code: &str| {
        let folded = code.to_lowercase();
        has_selector |= if local {
            folded.contains('[')
        } else {
            folded.contains(&table)
        };
        has_column |=
            folded.contains(&column) || folded.replace('\'', "").contains(&unescaped_column);
        has_selector && has_column
    };
    if !input.contains('"') {
        return candidate(input);
    }
    let mut start = 0;
    let mut brackets = 0_usize;
    let mut quoted_sheet = false;
    let mut string = false;
    let mut chars = input.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '"' if brackets == 0 && !quoted_sheet => {
                if string {
                    if chars.next_if(|&(_, next)| next == '"').is_some() {
                        continue;
                    }
                    start = index + 1;
                } else if candidate(&input[start..index]) {
                    return true;
                }
                string = !string;
            }
            _ if string => {}
            '\'' if brackets > 0 => {
                chars.next();
            }
            '\'' => quoted_sheet = !quoted_sheet,
            _ if quoted_sheet => {}
            '[' => brackets += 1,
            ']' => brackets = brackets.saturating_sub(1),
            _ => {}
        }
    }
    !string && candidate(&input[start..])
}

#[doc(hidden)]
pub fn rewrite_table_columns<'a>(
    input: &str,
    caption: impl Fn(Option<&str>, &str, &str) -> Option<&'a str>,
) -> Result<Option<String>, ParseError> {
    let mut lexer = Lexer {
        input,
        pos: 0,
        table_refs: Some(Vec::new()),
        is_range_name: &is_formula_parameter,
        is_sheet_span: None,
        evaluation_error: None,
    };
    let tokens = lexer.run()?;
    let table_refs = lexer.table_refs.take().unwrap();
    let mut spans = Vec::new();
    let mut indices = 0..tokens.len();
    for TableRefRewrite {
        start,
        selector_start,
        columns,
    } in table_refs
    {
        let Some(index) = indices.find(|&index| tokens[index].start == start) else {
            continue;
        };
        let token = &tokens[index];
        let TokKind::TableRef { sheet, table, .. } = &token.kind else {
            continue;
        };
        for (column, span) in columns {
            if let Some(new) = caption(sheet.as_deref(), table, &column) {
                let prefix = &input[selector_start + 1..span.start];
                let needs_brackets = new
                    .chars()
                    .any(|ch| !(ch.is_alphabetic() || ch.is_ascii_digit() || ch == '_'));
                let bare = matches!(prefix, "" | "@");
                let span =
                    if !needs_brackets && matches!(prefix, "[" | "@[") && token.end == span.end + 2
                    {
                        span.start - 1..span.end + 1
                    } else {
                        span
                    };
                spans.push((span, new, bare && needs_brackets));
            }
        }
    }
    if spans.is_empty() {
        return Ok(None);
    }
    ensure_unambiguous_spaced_calls(input, &tokens)?;
    let mut chars = input.chars().count();
    for (span, new, bracket) in &spans {
        if !new.chars().all(|ch| {
            matches!(ch, '\t' | '\n' | '\r') || (ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}')
        }) {
            return Err(ParseError::new(
                span.start,
                "the new caption contains a character XML cannot carry in a formula",
            ));
        }
        chars -= input[span.clone()].chars().count();
        chars += new.chars().count()
            + new.matches(TABLE_ESCAPE_CHARS).count()
            + 2 * usize::from(*bracket);
    }
    if chars > MAX_FORMULA_CHARS {
        return Err(ParseError::new(
            0,
            format!("formula exceeds Excel's {MAX_FORMULA_CHARS}-character limit"),
        ));
    }
    let mut result = String::with_capacity(input.len());
    let mut cursor = 0;
    for (span, new, bracket) in spans {
        result.push_str(&input[cursor..span.start]);
        if bracket {
            result.push('[');
        }
        for ch in new.chars() {
            if TABLE_ESCAPE_CHARS.contains(&ch) {
                result.push('\'');
            }
            result.push(ch);
        }
        if bracket {
            result.push(']');
        }
        cursor = span.end;
    }
    result.push_str(&input[cursor..]);
    Ok(Some(result))
}

fn workbook_scope(scope: &str) -> Option<&str> {
    scope
        .split_once('[')
        .and_then(|(_, rest)| rest.split_once(']'))
        .filter(|(book, _)| !book.is_empty())
        .map(|(_, sheet)| sheet)
}

fn is_word_char(ch: char) -> bool {
    ch.is_alphanumeric() || matches!(ch, '_' | '.' | '$' | '\\')
}

fn has_unquoted_sheet_chars(word: &str) -> bool {
    word.starts_with(|ch: char| ch.is_alphabetic() || ch == '_') && !word.contains(['$', '\\'])
}

fn is_reference(kind: &TokKind) -> bool {
    matches!(
        kind,
        TokKind::Ref { .. }
            | TokKind::Range { .. }
            | TokKind::ColumnRange { .. }
            | TokKind::RowRange { .. }
            | TokKind::TableRef { .. }
            | TokKind::ThreeDRef { .. }
            | TokKind::ExternalRef(_)
            | TokKind::Name { .. }
            | TokKind::Ident(_)
            | TokKind::ErrLit(ErrorValue::Ref)
    )
}

fn is_intersection(left: &TokKind, right: &TokKind) -> bool {
    (is_reference(left) || *left == TokKind::RParen)
        && (is_reference(right) || *right == TokKind::LParen)
        && !(*right == TokKind::LParen && matches!(left, TokKind::Ident(_) | TokKind::Name { .. }))
}

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
    table_refs: Option<Vec<TableRefRewrite>>,
    is_range_name: &'a dyn Fn(&str) -> bool,
    is_sheet_span: Option<&'a SheetSpan<'a>>,
    evaluation_error: Option<&'a mut Option<ParseError>>,
}

struct TableRefRewrite {
    start: usize,
    selector_start: usize,
    columns: Vec<(String, std::ops::Range<usize>)>,
}

impl Lexer<'_> {
    fn run(&mut self) -> Result<Vec<Token>, ParseError> {
        if self.input.len() > MAX_FORMULA_BYTES {
            return Err(ParseError::new(
                MAX_FORMULA_BYTES,
                "formula length exceeded cap",
            ));
        }
        let mut out = Vec::new();
        loop {
            let whitespace = self.pos;
            self.skip_ws();
            let start = self.pos;
            let Some(c) = self.peek() else { break };
            let kind = match c {
                '0'..='9' => self.lex_number_or_row_range()?,
                '.' => self.lex_number()?,
                '"' => self.lex_string()?,
                '#' => self.lex_error_literal()?,
                '\'' => self.lex_quoted_sheet_ref()?,
                '[' => self.lex_bracket_ref(start)?,
                '(' => self.punct(TokKind::LParen),
                ')' => self.punct(TokKind::RParen),
                ',' => self.punct(TokKind::Comma),
                ':' => self.punct(TokKind::Colon),
                '{' => self.punct(TokKind::LBrace),
                '}' => self.punct(TokKind::RBrace),
                ';' => self.punct(TokKind::Semicolon),
                '+' => self.punct(TokKind::Plus),
                '-' => self.punct(TokKind::Minus),
                '*' => self.punct(TokKind::Star),
                '/' => self.punct(TokKind::Slash),
                '^' => self.punct(TokKind::Caret),
                '&' => self.punct(TokKind::Amp),
                '%' => self.punct(TokKind::Percent),
                '=' => self.punct(TokKind::Eq),
                '<' => {
                    self.bump();
                    match self.peek() {
                        Some('=') => {
                            self.bump();
                            TokKind::Le
                        }
                        Some('>') => {
                            self.bump();
                            TokKind::Ne
                        }
                        _ => TokKind::Lt,
                    }
                }
                '>' => {
                    self.bump();
                    match self.peek() {
                        Some('=') => {
                            self.bump();
                            TokKind::Ge
                        }
                        _ => TokKind::Gt,
                    }
                }
                c if c.is_alphabetic() || matches!(c, '$' | '_' | '\\') => {
                    self.lex_word_or_ref()?
                }
                other => {
                    return Err(ParseError::new(
                        start,
                        format!("unexpected character {other:?}"),
                    ));
                }
            };
            if self.input[whitespace..start].contains(' ')
                && out
                    .last()
                    .is_some_and(|token: &Token| is_intersection(&token.kind, &kind))
            {
                out.push(Token {
                    kind: TokKind::Intersection,
                    start: whitespace,
                    end: start,
                });
            }
            out.push(Token {
                kind,
                start,
                end: self.pos,
            });
            if out.len() > MAX_TOKENS {
                return Err(ParseError::new(start, "formula token count exceeded cap"));
            }
        }
        Ok(out)
    }

    fn punct(&mut self, kind: TokKind) -> TokKind {
        self.bump();
        kind
    }

    fn peek(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn bump(&mut self) -> Option<char> {
        let c = self.peek()?;
        self.pos += c.len_utf8();
        Some(c)
    }

    fn skip_ws(&mut self) {
        self.pos = self.input.len() - self.after_ws().len();
    }

    fn after_ws(&self) -> &str {
        self.input[self.pos..].trim_start_matches([' ', '\t', '\r', '\n'])
    }

    fn consume_bang(&mut self) -> bool {
        if !self.after_ws().starts_with('!') {
            return false;
        }
        let qualifier_end = self.pos;
        self.skip_ws();
        let bang = self.pos;
        self.bump();
        self.skip_ws();
        if bang != qualifier_end || self.pos != bang + 1 {
            self.reject_for_evaluation(bang, "whitespace around a sheet qualifier is unsupported");
        }
        true
    }

    fn reject_for_evaluation(&mut self, pos: usize, message: &'static str) {
        if let Some(error) = &mut self.evaluation_error {
            error.get_or_insert_with(|| ParseError::new(pos, message));
        }
    }

    /// read a maximal name/reference run.
    fn read_word(&mut self) -> String {
        let start = self.pos;
        while self.peek().is_some_and(is_word_char) {
            self.bump();
        }
        self.input[start..self.pos].to_string()
    }

    fn lex_number(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        while self.peek().is_some_and(|c| c.is_ascii_digit()) {
            self.bump();
        }
        if self.peek() == Some('.') {
            self.bump();
            while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
            }
        }
        if matches!(self.peek(), Some('e' | 'E')) {
            self.bump();
            if matches!(self.peek(), Some('+' | '-')) {
                self.bump();
            }
            let exp_start = self.pos;
            while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
            }
            if self.pos == exp_start {
                return Err(ParseError::new(start, "malformed number exponent"));
            }
        }
        let text = &self.input[start..self.pos];
        let value = literal_value(text)
            .ok_or_else(|| ParseError::new(start, format!("malformed number {text:?}")))?;
        let value = literal_in_range(value)
            .ok_or_else(|| ParseError::new(start, "number is outside the range excel accepts"))?;
        Ok(TokKind::Num(value))
    }

    fn lex_number_or_row_range(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        let rest = &self.input[start..];
        let after_digits = rest.trim_start_matches(|ch: char| ch.is_ascii_digit());
        if after_digits.trim_start().starts_with(':') {
            self.pos = self.input.len() - after_digits.len();
            let first = &self.input[start..self.pos];
            if let Some(kind) = self.try_range_token(&None, first) {
                return Ok(kind);
            }
            self.pos = start;
        }
        self.lex_number()
    }

    fn lex_string(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        self.bump();
        let mut s = String::new();
        loop {
            match self.bump() {
                None => return Err(ParseError::new(start, "unterminated string")),
                Some('"') => {
                    if self.peek() == Some('"') {
                        self.bump();
                        s.push('"');
                    } else {
                        break;
                    }
                }
                Some(c) => s.push(c),
            }
        }
        Ok(TokKind::Str(s))
    }

    fn lex_error_literal(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        let rest = &self.input[start..];
        for (lit, val) in ERROR_LITERALS {
            if rest.starts_with(lit) {
                self.pos += lit.len();
                return Ok(TokKind::ErrLit(*val));
            }
        }
        Err(ParseError::new(start, "unknown error literal"))
    }

    fn lex_bracket_ref(&mut self, start: usize) -> Result<TokKind, ParseError> {
        if let Some(kind) = self.try_external_ref(start)? {
            return Ok(kind);
        }
        let kind = self.lex_table_ref(None, String::new(), start)?;
        if matches!(&self.input[start..self.pos], "[,]" | "[:]")
            || matches!(&kind, TokKind::TableRef { spec, .. } if *spec == TableSpec::default())
        {
            return Err(ParseError::new(start, "empty local structured reference"));
        }
        if self
            .peek()
            .is_some_and(|ch| matches!(ch, '[' | '\'') || is_word_char(ch))
        {
            return Err(ParseError::new(
                self.pos,
                "expected an operator after the selector",
            ));
        }
        Ok(kind)
    }

    fn try_external_ref(&mut self, start: usize) -> Result<Option<TokKind>, ParseError> {
        let input = self.input;
        let bracket = self.pos;
        let Some(book) = input[bracket..].strip_prefix('[') else {
            return Ok(None);
        };
        let Some(end) = book.find(['[', ']', '"', '\'']) else {
            return Ok(None);
        };
        if end == 0
            || !book[end..].starts_with(']')
            || TableBand::parse(&book[..end]).is_some()
            || book[..end].contains(|ch: char| {
                ch.is_whitespace()
                    || matches!(ch, '@' | ':' | '+' | '*' | '/' | '\\' | '<' | '>' | '=')
            })
        {
            return Ok(None);
        }
        let prefix_end = bracket + end + 2;
        self.pos = prefix_end;
        let reference = if self.peek() == Some('!') && self.consume_bang() {
            if self.input[self.pos..].starts_with("#REF!") {
                let error = self.lex_error_literal()?;
                return Ok(Some(TokKind::ExternalRef(Box::new(error))));
            }
            let name_start = self.pos;
            let name = self.read_word();
            if name.is_empty() {
                return Err(ParseError::new(
                    self.pos,
                    "expected an external defined name",
                ));
            }
            let index = &book[..end];
            if bracket == start
                && index.bytes().all(|byte| byte.is_ascii_digit())
                && &input[prefix_end..name_start] == "!"
                && !name.contains(['$', '\\'])
                && self.peek() == Some('(')
            {
                let prefix = &input[start..prefix_end];
                return Ok(Some(TokKind::Ident(format!("{prefix}!{name}"))));
            }
            if self.peek() == Some('[') {
                self.lex_table_ref(None, name, start)?
            } else {
                TokKind::Name { scope: None, name }
            }
        } else {
            let sheet = self.read_word();
            if sheet.is_empty() {
                self.pos = bracket;
                return Ok(None);
            }
            let Some(kind) = self.try_sheet_reference(&sheet, false, start)? else {
                self.pos = bracket;
                return Ok(None);
            };
            kind
        };
        Ok(Some(TokKind::ExternalRef(Box::new(reference))))
    }

    fn read_sheet_endpoint(&mut self) -> Option<String> {
        if self.peek() == Some('\'') {
            let token = parse_sheet_token(self.input, self.pos)?;
            self.pos = token.end;
            Some(token.name)
        } else {
            let word = self.read_word();
            (!word.is_empty()).then_some(word)
        }
    }

    fn lex_quoted_sheet_ref(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        let sheet = self
            .read_sheet_endpoint()
            .ok_or_else(|| ParseError::new(start, "unterminated sheet name"))?;
        let kind = self
            .try_sheet_reference(&sheet, true, start)?
            .ok_or_else(|| ParseError::new(self.pos, "expected a sheet qualifier"))?;
        if is_path_qualifier(&sheet) {
            self.reject_for_evaluation(start, "workbook path qualifiers are unsupported");
        }
        if (workbook_scope(&sheet).is_some() || is_path_qualifier(&sheet))
            && !(matches!(kind, TokKind::Name { .. }) && self.after_ws().starts_with('('))
        {
            return Ok(TokKind::ExternalRef(Box::new(kind)));
        }
        Ok(kind)
    }

    fn try_sheet_reference(
        &mut self,
        first: &str,
        quoted: bool,
        start: usize,
    ) -> Result<Option<TokKind>, ParseError> {
        if self.consume_bang() {
            if quoted
                && let Some((first, last)) = workbook_scope(first)
                    .or_else(|| (!is_path_qualifier(first)).then_some(first))
                    .and_then(|scope| scope.split_once(':'))
            {
                if first.is_empty() || last.is_empty() || last.contains(':') {
                    return Err(ParseError::new(start, "invalid 3-D sheet qualifier"));
                }
                return self.read_3d_target(first, last).map(Some).ok_or_else(|| {
                    ParseError::new(start, "expected an address after a 3-D qualifier")
                });
            }
            if !quoted && !has_unquoted_sheet_chars(first) {
                return Err(ParseError::new(start, "sheet name must be quoted"));
            }
            return self.lex_reference(first.to_owned(), start).map(Some);
        }
        if self.peek() != Some(':')
            || !(quoted || is_unquoted_sheet(first))
            || (!quoted && start + first.len() == self.pos && (self.is_range_name)(first))
        {
            return Ok(None);
        }
        let save = self.pos;
        self.bump();
        let quoted_last = self.peek() == Some('\'');
        if let Some(last) = self.read_sheet_endpoint()
            && self.consume_bang()
        {
            let first = workbook_scope(first).unwrap_or(first);
            if (quoted || quoted_last)
                && (first.is_empty()
                    || last.is_empty()
                    || first.contains(':')
                    || last.contains(':'))
            {
                return Err(ParseError::new(start, "invalid 3-D sheet qualifier"));
            }
            if (quoted_last || has_unquoted_sheet_chars(&last))
                && let Some(kind) = self.read_3d_target(first, &last)
                && (quoted
                    || &self.input[start..save] != first
                    || last.contains([':', '['])
                    || is_path_qualifier(&last)
                    || self.is_sheet_span.is_none_or(|is_span| is_span(first, &last)))
            {
                return Ok(Some(kind));
            }
        }
        self.pos = save;
        Ok(None)
    }

    fn read_3d_target(&mut self, first: &str, last: &str) -> Option<TokKind> {
        let target = if self.input[self.pos..].starts_with("#REF!") {
            self.lex_error_literal().ok()?
        } else {
            let word = self.read_word();
            self.try_address(&None, &word)?
        };
        Some(TokKind::ThreeDRef {
            first: first.to_owned(),
            last: last.to_owned(),
            target: Box::new(target),
        })
    }

    fn lex_word_or_ref(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        let word = self.read_word();
        let end = self.pos;
        if word.len() == 1
            && word.as_bytes()[0].is_ascii_alphabetic()
            && let Some(path) = self.input[end..].strip_prefix(":\\")
        {
            let after_path = path.trim_start_matches(is_word_char);
            self.pos = self.input.len() - after_path.len();
        }
        let path_end = self.pos;
        if (path_end != end || word.ends_with('\\'))
            && let Some(kind) = self.try_external_ref(start)?
        {
            if !self.input[start..path_end].ends_with('\\') {
                return Err(ParseError::new(
                    path_end,
                    "expected a path separator before the workbook",
                ));
            }
            return Ok(kind);
        }
        self.pos = end;
        if let Some(kind) = self.try_sheet_reference(&word, false, start)? {
            return Ok(kind);
        }
        if self.peek() == Some('[') {
            return self.lex_table_ref(None, word, start);
        }
        if self.peek() == Some('(') {
            return Ok(TokKind::Ident(word));
        }
        if word.eq_ignore_ascii_case("TRUE") {
            return Ok(TokKind::Bool(true));
        }
        if word.eq_ignore_ascii_case("FALSE") {
            return Ok(TokKind::Bool(false));
        }
        Ok(self
            .try_address(&None, &word)
            .unwrap_or(TokKind::Ident(word)))
    }

    /// `Table[...]`: either one bare item (`Table[Col]`, `Table[#Headers]`,
    /// `Table[@Col]`) or a bracketed list (`Table[[#This Row],[Col]]`).
    fn lex_table_ref(
        &mut self,
        sheet: Option<String>,
        table: String,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        let selector_start = self.pos;
        let mut columns = self.table_refs.as_ref().map(|_| Vec::new());
        let spec = self.read_table_spec(start, |column, span| {
            if let Some(columns) = &mut columns {
                columns.push((column.to_owned(), span));
            }
        })?;
        if let (Some(table_refs), Some(columns)) = (&mut self.table_refs, columns) {
            table_refs.push(TableRefRewrite {
                start,
                selector_start,
                columns,
            });
        }
        Ok(TokKind::TableRef { sheet, table, spec })
    }

    fn read_table_spec(
        &mut self,
        start: usize,
        mut column: impl FnMut(&str, std::ops::Range<usize>),
    ) -> Result<TableSpec, ParseError> {
        self.bump();
        let mut spec = TableSpec::default();
        if self.peek() == Some(']') {
            self.bump();
            return Ok(spec);
        }
        let rest = &self.input[self.pos..];
        let nested = rest.trim_start_matches([' ', '\t', '\n', '\r']);
        if nested.starts_with("@[") {
            self.skip_ws();
            self.bump();
            spec.bands.push(TableBand::ThisRow);
        } else if !nested.starts_with('[') {
            self.read_table_spec_item(&mut spec, start, &mut column)?;
            return Ok(spec);
        }
        let mut items = 0;
        loop {
            self.skip_ws();
            match self.peek() {
                Some('[') => {
                    self.bump();
                    items += 1;
                    if items > MAX_TABLE_SPEC_ITEMS {
                        return Err(ParseError::new(
                            start,
                            "table reference item count exceeded cap",
                        ));
                    }
                    self.read_table_spec_item(&mut spec, start, &mut column)?;
                }
                Some(',' | ':') => {
                    self.bump();
                }
                Some(']') => {
                    self.bump();
                    break;
                }
                _ => return Err(ParseError::new(start, "malformed table reference")),
            }
        }
        Ok(spec)
    }

    fn read_table_spec_item(
        &mut self,
        spec: &mut TableSpec,
        start: usize,
        column: &mut impl FnMut(&str, std::ops::Range<usize>),
    ) -> Result<(), ParseError> {
        if self.peek() == Some('@') {
            self.bump();
            spec.bands.push(TableBand::ThisRow);
        }
        let item_start = self.pos;
        let is_band = self.peek() == Some('#');
        let mut item = String::new();
        loop {
            match self.bump() {
                None => return Err(ParseError::new(start, "unterminated table reference")),
                Some(']') => break,
                Some('\'') => match self.bump() {
                    None => return Err(ParseError::new(start, "unterminated table reference")),
                    Some(escaped) => item.push(escaped),
                },
                Some(c) => item.push(c),
            }
        }
        if item.is_empty() {
            return Ok(());
        }
        if is_band {
            let band = TableBand::parse(&item)
                .ok_or_else(|| ParseError::new(start, "unknown table band selector"))?;
            spec.bands.push(band);
        } else {
            column(&item, item_start..self.pos - 1);
            if spec.first_column.is_none() {
                spec.first_column = Some(item);
            } else if spec.last_column.is_none() {
                spec.last_column = Some(item);
            } else {
                return Err(ParseError::new(start, "too many table columns"));
            }
        }
        Ok(())
    }

    fn lex_reference(&mut self, sheet: String, start: usize) -> Result<TokKind, ParseError> {
        if self.input[self.pos..].starts_with("#REF!") {
            return self.lex_error_literal();
        }
        let word = self.read_word();
        if word.is_empty() {
            return Err(ParseError::new(start, "expected reference or defined name"));
        }
        let sheet = Some(sheet);
        if self.peek() == Some('[') {
            return self.lex_table_ref(sheet, word, start);
        }
        if self.evaluation_error.is_some()
            && self.peek() == Some('(')
            && CellRef::parse_a1(&word.to_ascii_uppercase()).is_ok()
        {
            self.reject_for_evaluation(start, "qualified cell-shaped calls are unsupported");
        }
        Ok(self.try_address(&sheet, &word).unwrap_or(TokKind::Name {
            scope: sheet,
            name: word,
        }))
    }

    fn try_address(&mut self, sheet: &Option<String>, word: &str) -> Option<TokKind> {
        if let Some(range) = self.try_range_token(sheet, word) {
            return Some(range);
        }
        if !self.ends_address() {
            return None;
        }
        let cell = CellRef::parse_a1(&word.to_ascii_uppercase()).ok()?;
        Some(TokKind::Ref {
            sheet: sheet.clone(),
            cell,
        })
    }

    fn ends_address(&self) -> bool {
        !matches!(self.peek(), Some('[' | '(')) && !self.after_ws().starts_with('!')
    }

    fn try_range_token(&mut self, sheet: &Option<String>, word: &str) -> Option<TokKind> {
        if !self.after_ws().starts_with(':') {
            return None;
        }
        let save = self.pos;
        self.skip_ws();
        self.bump();
        self.skip_ws();
        let adjacent_colon = self.pos == save + 1;
        let end = self.read_word();
        let ends_address = self.ends_address();
        if ends_address || self.evaluation_error.is_some() {
            let source = format!("{word}:{end}");
            let kind = CellRange::parse_a1(&source.to_ascii_uppercase())
                .map(|range| TokKind::Range {
                    sheet: sheet.clone(),
                    range,
                })
                .or_else(|_| {
                    ColumnRange::parse_a1(&source).map(|range| TokKind::ColumnRange {
                        sheet: sheet.clone(),
                        range,
                    })
                })
                .or_else(|_| {
                    RowRange::parse_a1(&source).map(|range| TokKind::RowRange {
                        sheet: sheet.clone(),
                        range,
                    })
                });
            if let Ok(kind) = kind {
                if ends_address {
                    return Some(kind);
                }
                if adjacent_colon || !matches!(kind, TokKind::Range { .. }) {
                    self.reject_for_evaluation(save, "range endpoint is unsupported");
                }
            }
        }
        self.pos = save;
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(src: &str) -> Vec<TokKind> {
        lex(src).unwrap().into_iter().map(|t| t.kind).collect()
    }

    #[test]
    fn rewrites_only_the_matching_structured_reference_columns() {
        let source =
            r#"T[Old]+t[[#This Row], [old]]+T[[Old]:[X]]+T[@OLD]+[@Old]+Other[Old]+"T[Old]""#;
        let expected = concat!(
            "T[['#New']]]+t[[#This Row], ['#New']]]+T[['#New']]:[X]]",
            r#"+T[@['#New']]]+[@['#New']]]+Other[Old]+"T[Old]""#,
        );
        let rewritten = rewrite_table_column(source, "T", "Data", true, "Old", "#New]")
            .unwrap()
            .unwrap();
        assert_eq!(rewritten, expected);
        assert_eq!(
            rewrite_table_column(&rewritten, "T", "Data", true, "#New]", "New")
                .unwrap()
                .unwrap(),
            r#"T[New]+t[[#This Row], [New]]+T[[New]:[X]]+T[@New]+[@New]+Other[Old]+"T[Old]""#
        );
    }

    #[test]
    fn rewritten_captions_with_punctuation_or_spaces_use_bracketed_selectors() {
        let source = "T[Old]+[@Old]+T[[#This Row],[Old]]";
        for caption in ["Margin(%)", "a b", "x-y"] {
            let rewritten = rewrite_table_column(source, "T", "Data", true, "Old", caption)
                .unwrap()
                .unwrap();
            assert_eq!(
                rewritten,
                format!("T[[{caption}]]+[@[{caption}]]+T[[#This Row],[{caption}]]")
            );
        }
    }

    #[test]
    fn a_local_selector_can_follow_an_internal_name_or_be_an_external_function_argument() {
        let source = "SUM(Data!AllowedRange [@Old])+[1]!ExternalFn([@Old])";
        assert_eq!(
            rewrite_table_column(source, "T", "Data", true, "Old", "New")
                .unwrap()
                .as_deref(),
            Some("SUM(Data!AllowedRange [@New])+[1]!ExternalFn([@New])")
        );
    }

    #[test]
    fn table_column_candidates_skip_strings_with_escaped_quotes() {
        let literal = r#"SUM(Jan:Mar!A1)+N("Table1[""Status""]")"#;
        for local in [false, true] {
            assert!(!may_reference_table_column(
                literal, "Table1", local, "Status"
            ));
        }
        let source = r#"'Data "Q"'!Table1["Status"]+N("Table1[""Status""]")"#;
        assert_eq!(
            rewrite_table_column(source, "Table1", "Data \"Q\"", false, "\"Status\"", "State")
                .unwrap()
                .as_deref(),
            Some(r#"'Data "Q"'!Table1[State]+N("Table1[""Status""]")"#)
        );
    }

    #[test]
    fn the_rewrite_lexer_accepts_3d_qualifiers() {
        for (source, scope) in [
            ("Jan:Mar!A1", "Jan:Mar"),
            ("'Jan:Mar'!A1", "Jan:Mar"),
            ("Jan:'End Sheet'!A1", "Jan:End Sheet"),
            ("'Start Sheet':Mar!#REF!", "Start Sheet:Mar"),
            ("Sheet1:Sheet2!A1", "Sheet1:Sheet2"),
            ("'Sheet1:Sheet2'!A1", "Sheet1:Sheet2"),
            ("Sheet1:Sheet2!$A$1", "Sheet1:Sheet2"),
            ("Sheet1:Sheet2!A1:B2", "Sheet1:Sheet2"),
            ("Sheet1:Sheet2!A:A", "Sheet1:Sheet2"),
            ("Sheet1:Sheet2!1:1", "Sheet1:Sheet2"),
            ("Sheet1:Sheet2!#REF!", "Sheet1:Sheet2"),
            ("Q1_Sales:Data_1!A1", "Q1_Sales:Data_1"),
            ("Résumé:Sheet2!A1", "Résumé:Sheet2"),
        ] {
            let tokens = lex_for_rewrite(source).unwrap();
            assert_eq!(tokens.len(), 1, "{source}");
            let (first, last) = scope.split_once(':').unwrap();
            assert!(
                matches!(&tokens[0].kind,
                TokKind::ThreeDRef { first: start, last: end, target }
                    if start == first && end == last
                        && matches!(target.as_ref(), TokKind::Ref { .. } | TokKind::Range { .. }
                            | TokKind::ColumnRange { .. } | TokKind::RowRange { .. }
                            | TokKind::ErrLit(ErrorValue::Ref))),
                "{source}"
            );
            assert_eq!((tokens[0].start, tokens[0].end), (0, source.len()));
        }
    }

    #[test]
    fn the_rewrite_lexer_keeps_the_range_operator_before_qualified_names() {
        for (source, first, sheet, name) in [
            ("Name:Sheet!Name", "Name", "Sheet", "Name"),
            ("Sheet1:Sheet2!Total", "Sheet1", "Sheet2", "Total"),
        ] {
            let tokens = lex_for_rewrite(source).unwrap();
            assert_eq!(
                tokens
                    .into_iter()
                    .map(|token| token.kind)
                    .collect::<Vec<_>>(),
                [
                    TokKind::Ident(first.into()),
                    TokKind::Colon,
                    TokKind::Name {
                        scope: Some(sheet.into()),
                        name: name.into(),
                    },
                ],
                "{source}"
            );
        }
    }

    #[test]
    fn table_ref_tokens_expose_the_table_scope_and_spec() {
        let tokens = lex("Sales[Amount]").unwrap();
        let TokKind::TableRef {
            sheet: None,
            table,
            spec,
        } = &tokens[0].kind
        else {
            panic!("expected a table reference");
        };
        assert_eq!(table, "Sales");
        assert_eq!(spec.first_column.as_deref(), Some("Amount"));
        assert!(matches!(
            &kinds("Sheet1!Sales[Amount]")[0],
            TokKind::TableRef { sheet: Some(sheet), table, .. }
                if sheet == "Sheet1" && table == "Sales"
        ));
    }

    #[test]
    fn a_qualified_table_reference_keeps_its_qualifier_and_table_identity() {
        let source = "Sheet2!Other[Old]+'Sales Data'!T[Old]+Sheet2!T[Old]+[1]!T[Old]+[@Old]";
        assert_eq!(
            rewrite_table_column(source, "T", "Sales Data", true, "Old", "New")
                .unwrap()
                .as_deref(),
            Some("Sheet2!Other[Old]+'Sales Data'!T[New]+Sheet2!T[Old]+[1]!T[Old]+[@New]")
        );
        assert_eq!(
            rewrite_table_column("Sheet2!T[Old]", "T", "Sheet2", false, "Old", "New")
                .unwrap()
                .as_deref(),
            Some("Sheet2!T[New]")
        );
    }

    #[test]
    fn the_rewrite_lexer_reads_external_workbook_prefixes() {
        for (source, reference) in [
            ("[1]Data!$A$1+Sales[Amt]", "[1]Data!$A$1"),
            ("Sales[[#This Row],[Amt]]*[12]Rates!B2", "[12]Rates!B2"),
            ("[3]!Total+1", "[3]!Total"),
            ("[Book.xlsx]Data!A1", "[Book.xlsx]Data!A1"),
        ] {
            let tokens = lex_for_rewrite(source).unwrap();
            let external = tokens
                .iter()
                .find(|token| matches!(token.kind, TokKind::ExternalRef(_)))
                .unwrap();
            assert_eq!(&source[external.start..external.end], reference);
        }
        assert!(matches!(
            lex_for_rewrite("Sales[Amt]").unwrap()[0].kind,
            TokKind::TableRef { .. }
        ));
        assert!(lex_for_rewrite("[]Data!A1").is_err());
    }

    #[test]
    fn lexer_modes_share_token_spans_and_kinds() {
        for source in [
            "SUM(Sales:Q2!#REF!)",
            "SUM(Sheet1:Sheet2!A1,Sheet1:Sheet2!#REF!,R2C3x:Sheet2!#REF!)",
            "SUM(Sheet1:'Sheet 2'!1:2,'Sheet 1':Sheet2!#REF!,_xlpm.start:Sheet2!A1)",
            "Amt:Tax[Old]",
            "SUM(Amt : Tax[Old])",
            "SUM(Sheet1!Amt:Tax[Old])",
            "SUM(Amt:Tax!Total)",
            "SUM(Amt:Tax!Table1[Old])",
            "SUM(Sheet1!Amt:Tax!B5)",
            "SUM(Amt : Tax!B5)",
            "[Book.xlsx]Sheet1:Sheet2!A1",
            "[#2025.xlsx]Sheet1:Sheet2!A1+[@Old]",
            "Amounts:Other!B5",
            "Amounts:Other!A1:B5",
            "Amounts:Other!$A:$A",
            "Amounts:Other!1:3",
            r"\Table1:Other!A1",
            "SUM(Amt:IF(C1,B1,B2))",
            "SUM(Amt:LET(x,B1,x))",
            "SUM(Amt:LET (x,B1,x),Amt:Tax (1:2),A1:B2 (C1:D2))",
            "SUM(A:N (1:2),A:T (1:2),Amt:MAX (1:2),Amt:IF (C1,B1,B2),LOG10 (1))",
            "A1!B2+TRUE!A1+R2C3!A1",
            "SUM([2024] Prices[Amt])",
            "SUM([1] Prices)",
            "SUM([1]Sheet1!A1 [@Old])",
            "SUM('[1]Sheet 1'!A1:B5 [@Old])",
            "SUM([1]!ExternalRange [@Old])",
            "SUM([Book.xlsx]Sheet1!A1 [@Old])",
            "SUM([1]Sheet1!ExternalName [[Old]:[X]])",
            "Sheet1 !A1+Table1[Old]",
            "Sheet1! A1+Table1[Old]",
            "'My Sheet' ! A1+Table1[Old]",
            "Sheet1:Sheet2 ! A1+Table1[Old]",
            "[@Qty]*Prices[UnitPrice]",
            "SUM([@Col]:[@Col2])+Prices[Amt]",
            "[#This Row]+[Col]+Table1[Col]+Sheet1!Table1[Col]",
            "'Sheet 1'!Table1[Col]+[1]Sheet1!Table1[Col]+'[1]Sheet 1'!Table1[Col]",
            "[1]Sheet1!A1+[1]!Name+[1]Sheet1:Sheet2!A1+[1]Sheet1:Sheet2!#REF!",
            "[1]!ExternalFn([@Old])",
            r"'C:\x\[Book.xlsx]Sheet1'!A1+C:\dir\[Book.xlsx]Sheet1!A1",
            "'https://example.com:8443/book.xlsx'!A1+T[Old]",
            "'It''s'!A1+'a!b'!A1+'It''s':'Other'!A1+'It''s:Other'!A1",
            "A1:Sheet2!B2+$A$1:Sheet2!B2:C3+LOG10:Sheet2!B2",
            "B2!A1:B2!C3+A1:B2[Col]",
            "A1:LOG10(1)+Sheet1!LOG10(1)+Sheet1:Sheet2!LOG10(1)",
            "Sheet1!A1 : B2+Sheet1:Sheet2!A1 : B2",
            r#"INDIRECT("R1C1",FALSE)+INDIRECT("R[-1]C[2]:R[1]C[4]",FALSE)"#,
            "R1C1+A1B+XFD1048577+'R1C1'!A1+R1C1:Other!A1",
            "R1C1:R2C2+R[-1]C[2]",
            r#"TRUE+FALSE+TRUE()+SUM(#REF!,#N/A,#DIV/0!,"[Text]:! ""quoted""")"#,
        ] {
            let expected = lex(source).unwrap_or_else(|error| panic!("{source}: {error}"));
            assert_eq!(lex_for_rewrite(source).unwrap(), expected, "{source}");
            let mut evaluation_error = None;
            assert_eq!(
                Lexer {
                    input: source,
                    pos: 0,
                    table_refs: Some(Vec::new()),
                    is_range_name: &is_formula_parameter,
                    is_sheet_span: None,
                    evaluation_error: Some(&mut evaluation_error),
                }
                .run()
                .unwrap(),
                expected,
                "{source}"
            );
        }
    }

    #[test]
    fn a_column_range_does_not_take_the_name_of_a_structured_reference() {
        for source in ["Amt:Tax[Old]", "Amt : Tax[Old]", "Sheet1!Amt:Tax[Old]"] {
            assert_eq!(
                rewrite_table_column(source, "LocalTbl", "Data", true, "Old", "New"),
                Ok(None)
            );
            assert_eq!(
                rewrite_table_column(source, "Tax", "Data", false, "Old", "New"),
                Ok(Some(source.replace("[Old]", "[New]")))
            );
        }
    }

    #[test]
    fn external_sheet_spans_do_not_capture_workbook_names_as_columns() {
        assert_eq!(
            rewrite_table_column(
                "SUM([Book.xlsx]Sheet1:Sheet2!A1)",
                "T",
                "Data",
                true,
                "Book.xlsx",
                "File",
            ),
            Ok(None)
        );
        assert_eq!(
            rewrite_table_column(
                "[#2025.xlsx]Sheet1:Sheet2!A1+[@Old]",
                "T",
                "Data",
                true,
                "Old",
                "New",
            ),
            Ok(Some("[#2025.xlsx]Sheet1:Sheet2!A1+[@New]".into()))
        );
    }

    #[test]
    fn short_names_before_a_qualified_table_do_not_block_header_rewrites() {
        for source in ["SUM(Amt:Tax!Table1[Old])", "SUM(Amt:Tax!Total)+Table1[Old]"] {
            assert_eq!(
                rewrite_table_column(source, "Table1", "Tax", false, "Old", "New"),
                Ok(Some(source.replace("[Old]", "[New]")))
            );
        }
    }

    #[test]
    fn intersection_with_an_external_reference_keeps_the_local_row_selector() {
        for reference in [
            "[1]Sheet1!A1",
            "'[1]Sheet 1'!A1:B5",
            "[1]!ExternalRange",
            "[Book.xlsx]Sheet1!A1",
        ] {
            let source = format!("SUM({reference} [@Old])");
            let tokens = lex(&source).unwrap();
            assert_eq!(tokens[3].kind, TokKind::Intersection, "{source}");
            assert_eq!(&source[tokens[3].start..tokens[3].end], " ");
            assert_eq!(
                rewrite_table_column(&source, "T", "Data", true, "Old", "New"),
                Ok(Some(format!("SUM({reference} [@New])")))
            );
        }
    }

    #[test]
    fn local_selectors_after_external_names_follow_a_header_rename() {
        let source = concat!(
            "SUM([1]Sheet1!ExternalName [Old])+",
            "SUM([1]Sheet1!ExternalName [[Old]:[X]])+",
            "SUM([1]Sheet1!ExternalName [@Old])+",
            "SUM(([1]Sheet1!ExternalName) [Old])+[1]Sheet1!T[Old]+[1]!T[Old]",
        );
        let expected = concat!(
            "SUM([1]Sheet1!ExternalName [New])+",
            "SUM([1]Sheet1!ExternalName [[New]:[X]])+",
            "SUM([1]Sheet1!ExternalName [@New])+",
            "SUM(([1]Sheet1!ExternalName) [New])+[1]Sheet1!T[Old]+[1]!T[Old]",
        );
        assert_eq!(
            rewrite_table_column(source, "T", "Data", true, "Old", "New"),
            Ok(Some(expected.into()))
        );
    }

    #[test]
    fn deleted_external_names_and_address_shaped_span_endpoints_keep_their_spans() {
        for reference in [
            "[1]!#REF!",
            "Sales:Q2!A1",
            "Sales:R1C1!#REF!",
            "Sales:TRUE!#REF!",
        ] {
            let source = format!("SUM({reference})+T[Old]");
            let tokens = lex(&source).unwrap();
            let token = &tokens[2];
            assert_eq!(&source[token.start..token.end], reference);
            assert!(matches!(
                token.kind,
                TokKind::ExternalRef(_) | TokKind::ThreeDRef { .. }
            ));
            assert_eq!(lex_for_rewrite(&source).unwrap(), tokens);
            assert_eq!(
                rewrite_table_column(&source, "T", "Data", false, "Old", "New"),
                Ok(Some(format!("SUM({reference})+T[New]")))
            );
        }
    }

    #[test]
    fn whitespace_distinguishes_builtin_calls_from_reference_intersections() {
        for name in ["LOG10", "IF"] {
            assert!(matches!(
                &kinds(&format!("{name}(1)"))[0], TokKind::Ident(word) if word == name
            ));
        }
        assert!(matches!(
            kinds("LOG10 (1)")[0],
            TokKind::Ref { sheet: None, .. }
        ));
        for columns in ["A:N", "A:T", "Amt:MAX", "Amt:IF", "Amt:LET"] {
            let source = format!("SUM({columns} (1:2))");
            let tokens = lex(&source).unwrap();
            assert!(
                matches!(tokens[2].kind, TokKind::ColumnRange { .. }),
                "{source}"
            );
            assert_eq!(&source[tokens[2].start..tokens[2].end], columns);
            assert_eq!(tokens[3].kind, TokKind::Intersection);
        }
        assert_eq!(kinds("A1 (B1)")[1], TokKind::Intersection);
        assert!(!kinds("A1\tB1").contains(&TokKind::Intersection));
        assert!(lex("[Col][Col2]").is_err());
    }

    #[test]
    fn ambiguous_spaced_calls_refuse_table_column_rewrites() {
        for reference in ["LOG10 (1)", "Amt:IF (C1,B1,B2)", "Amt:LET (x,B1,x)"] {
            let source = format!("SUM({reference})+T[Old]");
            let tokens = lex_for_rewrite(&source).unwrap();
            assert!(matches!(
                tokens[2].kind,
                TokKind::Ref { .. } | TokKind::ColumnRange { .. }
            ));
            assert_eq!(
                rewrite_table_column(&source, "T", "Data", false, "Old", "New"),
                Err(ParseError::new(4, "ambiguous spaced function call")),
                "{source}"
            );
        }
        assert_eq!(
            rewrite_table_column(
                "SUM(Amt:MAX (1:2))+T[Old]",
                "T",
                "Data",
                false,
                "Old",
                "New"
            ),
            Ok(Some("SUM(Amt:MAX (1:2))+T[New]".into())),
        );
    }

    #[test]
    fn quoted_workbook_paths_without_book_brackets_lex_as_external() {
        assert!(matches!(
            &kinds(r"'C:\Books\Data.xlsx'!A1")[..],
            [TokKind::ExternalRef(target)] if matches!(target.as_ref(), TokKind::Ref { .. })
        ));
        assert!(matches!(
            &kinds("'https://example.com/book.xlsx'!MyName")[..],
            [TokKind::ExternalRef(target)] if matches!(target.as_ref(), TokKind::Name { .. })
        ));
    }

    #[test]
    fn existing_unquoted_sheet_qualifiers_and_spans_keep_their_token_kinds() {
        for source in ["R1C1b!B2", "RCost!A1", "R2C3x!A1"] {
            assert!(matches!(
                &kinds(source)[..],
                [TokKind::Ref { sheet: Some(_), .. }]
            ));
        }
        assert!(matches!(
            &kinds("RCPT!A1:B5")[..],
            [TokKind::Range { sheet: Some(_), .. }]
        ));
        assert!(matches!(
            &kinds("RC_Data!Total")[..],
            [TokKind::Name { scope: Some(_), .. }]
        ));
        for first in ["Sheet1", "R2C3x"] {
            for target in ["A1", "#REF!"] {
                let source = format!("{first}:Sheet2!{target}");
                assert!(matches!(
                    &kinds(&source)[..],
                    [TokKind::ThreeDRef { first: start, last, .. }] if start == first && last == "Sheet2"
                ));
            }
        }
    }

    #[test]
    fn bracketed_captions_are_not_mistaken_for_external_workbooks() {
        for source in [
            "[@Old]Sheet1!A1",
            "[a + b]Sheet1!A1",
            "[#This Row]!Name",
            "[2024] !A1",
            "[Book.xlsx]'Sheet 1'!A1",
        ] {
            assert!(lex(source).is_err(), "{source}");
        }
    }

    #[test]
    fn bracketed_band_keywords_cannot_qualify_external_references() {
        for band in ["#All", "#Headers", "#Data", "#Totals", "#This Row"] {
            assert!(matches!(
                &kinds(&format!("[{band}]"))[0],
                TokKind::TableRef { .. }
            ));
            assert!(lex(&format!("[{band}]!Name")).is_err());
            assert!(lex(&format!("[{band}]Sheet1!A1")).is_err());
        }
        assert!(matches!(
            &kinds("[#2025.xlsx]!Name")[0],
            TokKind::ExternalRef(_)
        ));
    }

    #[test]
    fn empty_local_structured_references_require_a_table_name() {
        for selector in ["[]", "[[]]", "[,]", "[:]"] {
            let error = Err(ParseError::new(0, "empty local structured reference"));
            assert_eq!(lex(selector), error, "{selector}");
            assert_eq!(lex_for_rewrite(selector), error, "{selector}");
            assert_eq!(
                rewrite_table_column(
                    &format!("{selector}+T[Old]"),
                    "T",
                    "Data",
                    true,
                    "Old",
                    "New"
                ),
                Err(ParseError::new(0, "empty local structured reference")),
                "{selector}"
            );
        }
        assert_eq!(
            rewrite_table_column("T[]+T[Old]", "T", "Data", true, "Old", "New"),
            Ok(Some("T[]+T[New]".into())),
        );
    }

    #[test]
    fn a_drive_path_requires_a_separator_before_the_workbook() {
        for tail in ["[Book.xlsx]Sheet1!A1", "[1]!Name"] {
            let valid = format!(r"C:\Foo\{tail}");
            assert!(matches!(&kinds(&valid)[..], [TokKind::ExternalRef(_)]));
            let invalid = format!(r"C:\Foo{tail}");
            assert_eq!(
                lex_for_rewrite(&invalid),
                Err(ParseError::new(
                    6,
                    "expected a path separator before the workbook",
                )),
            );
        }
    }

    #[test]
    fn quoted_sheet_spans_require_two_nonempty_endpoints() {
        for source in ["':Other'!A1", "'Other:'!A1", "'A:B:C'!A1"] {
            assert!(lex(source).is_err(), "{source}");
        }
    }

    #[test]
    fn whitespace_at_a_sheet_qualifier_does_not_block_header_rewrites() {
        for reference in [
            "Sheet1 !A1",
            "Sheet1! A1",
            "'My Sheet' !A1",
            "'My Sheet'! A1",
            "Sheet1:Sheet2 !A1",
            "Sheet1:Sheet2! A1",
        ] {
            assert_eq!(
                rewrite_table_column(
                    &format!("{reference}+Table1[Old]"),
                    "Table1",
                    "Data",
                    false,
                    "Old",
                    "New"
                ),
                Ok(Some(format!("{reference}+Table1[New]")))
            );
        }
    }

    #[test]
    fn lexes_numbers_and_operators() {
        assert_eq!(
            kinds("1 + 2.5 * 3e2"),
            vec![
                TokKind::Num(1.0),
                TokKind::Plus,
                TokKind::Num(2.5),
                TokKind::Star,
                TokKind::Num(300.0),
            ]
        );
    }

    #[test]
    fn lexes_comparison_operators() {
        assert_eq!(
            kinds("<= <> >= < > ="),
            vec![
                TokKind::Le,
                TokKind::Ne,
                TokKind::Ge,
                TokKind::Lt,
                TokKind::Gt,
                TokKind::Eq
            ]
        );
    }

    #[test]
    fn lexes_strings_with_escaped_quotes() {
        assert_eq!(kinds(r#""a""b""#), vec![TokKind::Str("a\"b".into())]);
        assert_eq!(kinds(r#""""#), vec![TokKind::Str(String::new())]);
    }

    #[test]
    fn lexes_booleans_case_insensitively() {
        assert_eq!(
            kinds("TRUE false True"),
            vec![
                TokKind::Bool(true),
                TokKind::Bool(false),
                TokKind::Bool(true)
            ]
        );
    }

    #[test]
    fn lexes_error_literals() {
        assert_eq!(kinds("#REF!"), vec![TokKind::ErrLit(ErrorValue::Ref)]);
        assert_eq!(kinds("#N/A"), vec![TokKind::ErrLit(ErrorValue::NA)]);
        assert_eq!(kinds("#DIV/0!"), vec![TokKind::ErrLit(ErrorValue::Div0)]);
        assert_eq!(kinds("#CALC!"), vec![TokKind::ErrLit(ErrorValue::Calc)]);
    }

    #[test]
    fn lexes_cell_refs_and_ranges() {
        assert_eq!(
            kinds("A1"),
            vec![TokKind::Ref {
                sheet: None,
                cell: CellRef::parse_a1("A1").unwrap()
            }]
        );
        assert_eq!(
            kinds("$A$1"),
            vec![TokKind::Ref {
                sheet: None,
                cell: CellRef::parse_a1("$A$1").unwrap()
            }]
        );
        match &kinds("A1:B2")[0] {
            TokKind::Range { sheet: None, range } => assert_eq!(range.to_a1(), "A1:B2"),
            other => panic!("expected range, got {other:?}"),
        }
    }

    /// a literal `A1:B2` stays one token; anything else after the `:` leaves
    /// the colon for the parser to read as the range operator.
    #[test]
    fn lexes_the_range_operator_only_when_the_end_is_not_an_address() {
        assert!(matches!(kinds("A1:B2").as_slice(), [TokKind::Range { .. }]));
        assert!(matches!(
            kinds("A:A").as_slice(),
            [TokKind::ColumnRange { .. }]
        ));
        assert!(matches!(
            kinds("A1:INDEX(A1:A9,3)").as_slice(),
            [
                TokKind::Ref { .. },
                TokKind::Colon,
                TokKind::Ident(_),
                TokKind::LParen,
                ..
            ]
        ));
        assert!(matches!(
            kinds("Data!$J$2:INDEX(Data!$J:$J,4)").as_slice(),
            [TokKind::Ref { .. }, TokKind::Colon, TokKind::Ident(_), ..]
        ));
        assert!(matches!(
            kinds("A:OFFSET(A1,1,1)").as_slice(),
            [TokKind::Ident(_), TokKind::Colon, TokKind::Ident(_), ..]
        ));
        assert_eq!(kinds("A1:B2:C3").len(), 3);
    }

    #[test]
    fn lexes_unquoted_sheet_names_with_r1c1_prefixes() {
        for name in ["R1C1b", "RCost", "R2C3x"] {
            let source = format!("{name}!B2");
            assert!(
                matches!(
                    &kinds(&source)[..],
                    [TokKind::Ref { sheet: Some(sheet), cell }]
                        if sheet == name && *cell == CellRef::new(1, 1)
                ),
                "{source}"
            );
        }
    }

    #[test]
    fn lexes_sheet_qualified_refs() {
        match &kinds("Sheet1!A1")[0] {
            TokKind::Ref {
                sheet: Some(s),
                cell,
            } => {
                assert_eq!(s, "Sheet1");
                assert_eq!(cell.to_a1(), "A1");
            }
            other => panic!("expected sheet ref, got {other:?}"),
        }
        match &kinds("'My Sheet'!A1:B2")[0] {
            TokKind::Range {
                sheet: Some(s),
                range,
            } => {
                assert_eq!(s, "My Sheet");
                assert_eq!(range.to_a1(), "A1:B2");
            }
            other => panic!("expected quoted sheet range, got {other:?}"),
        }
    }

    #[test]
    fn function_name_wins_over_ref_shape() {
        assert_eq!(kinds("LOG10(")[0], TokKind::Ident("LOG10".into()));
        assert_eq!(kinds("SUM(")[0], TokKind::Ident("SUM".into()));
    }

    #[test]
    fn lexes_defined_names_and_scopes() {
        assert_eq!(kinds("Tax_Rate"), vec![TokKind::Ident("Tax_Rate".into())]);
        assert_eq!(kinds("\\Legacy"), vec![TokKind::Ident("\\Legacy".into())]);
        assert_eq!(
            kinds("'Input Sheet'!LocalRate"),
            vec![TokKind::Name {
                scope: Some("Input Sheet".into()),
                name: "LocalRate".into(),
            }]
        );
    }

    #[test]
    fn rejects_excessive_table_reference_items() {
        let items = "[#All],".repeat(MAX_TABLE_SPEC_ITEMS + 1);
        let error = lex(&format!("Sales[{items}[Amount]]")).unwrap_err();
        assert!(error.message.contains("item count"));
    }

    #[test]
    fn rejects_unterminated_string() {
        let err = lex("\"abc").unwrap_err();
        assert_eq!(err.pos, 0);
    }

    #[test]
    fn rejects_unexpected_char() {
        assert!(lex("1 ~ 2").is_err());
    }

    #[test]
    fn rejects_excessive_token_count() {
        let formula = "%".repeat(MAX_TOKENS + 1);
        let error = lex(&formula).unwrap_err();
        assert!(error.message.contains("token count"));
    }

    #[test]
    fn accepts_exact_token_limit() {
        assert_eq!(lex(&"%".repeat(MAX_TOKENS)).unwrap().len(), MAX_TOKENS);
    }

    #[test]
    fn rejects_excessive_formula_length() {
        let error = lex(&"x".repeat(MAX_FORMULA_BYTES + 1)).unwrap_err();
        assert!(error.message.contains("length"));
    }

    #[test]
    fn rejects_non_finite_number_literals() {
        assert!(lex("1e999").unwrap_err().message.contains("range"));
    }

    fn number(src: &str) -> f64 {
        match lex(src).unwrap()[0].kind {
            TokKind::Num(value) => value,
            ref other => panic!("{src} lexed as {other:?}"),
        }
    }

    /// Excel 365 for Mac result: ATANH(0.9999999999999999) is ATANH of
    /// 0.999999999999999, so the literal keeps fifteen significant digits.
    #[test]
    fn number_literals_keep_fifteen_significant_digits() {
        assert_eq!(number("0.9999999999999999"), 0.999_999_999_999_999);
        assert_eq!(number("9007199254740993"), 9_007_199_254_740_990.0);
        assert_eq!(number("123456789012345678E2"), 1.234_567_890_123_45e19);
        assert_eq!(
            number("0.000001234567890123456789"),
            1.234_567_890_123_45e-6
        );
        assert_eq!(number("1.00000000000000"), 1.0);
        assert_eq!(number("0.1"), 0.1);
    }

    /// Excel 365 for Mac result: 9.99999999999999E+307 is accepted, a literal
    /// of 1E+308 or more cannot be entered.
    #[test]
    fn number_literals_stop_below_1e308() {
        assert_eq!(number("9.99999999999999E+307"), 9.999_999_999_999_99e307);
        assert_eq!(number("9.9999999999999999E307"), 9.999_999_999_999_99e307);
        for src in ["1E308", "1e+308", "1.5E308", "10E307", "1E309"] {
            assert!(lex(src).is_err(), "{src}");
        }
    }

    /// Excel 365 for Mac result: 2.2250738585072E-308*1, 1E-308*1, 100E-310*1,
    /// 0.1E-307*1, 0.00001E-303*1, 9E-309*1 and 1E-309*1 are 0, Excel
    /// rewriting the literal to 0; formulas holding 1E-310, 10E307 or 0.1E309
    /// are refused, the limits going by the value; 0.1234567890123459 reads
    /// as 0.123456789012345.
    #[test]
    fn literal_limits_go_by_the_value() {
        for src in [
            "2.2250738585072E-308",
            "1E-308",
            "100E-310",
            "0.1E-307",
            "0.00001E-303",
            "9E-309",
            "1E-309",
            "0E-400",
        ] {
            assert_eq!(number(src).to_bits(), 0, "{src}");
        }
        assert_eq!(number("2.2250738585073E-308"), 2.225_073_858_507_3e-308);
        assert_eq!(number("0.1234567890123459"), 0.123_456_789_012_345);
        for src in ["1E-310", "10E307", "0.1E309"] {
            assert!(lex(src).is_err(), "{src}");
        }
    }
}
