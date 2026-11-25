//! formula lexer: turns source (without the leading `=`) into position-tagged
//! tokens, resolving reference/range/sheet disambiguation.

use std::fmt;

use xlsx_model::{CellRange, CellRef, ErrorValue};

use crate::reference::{MAX_TABLE_SPEC_ITEMS, column, row};
use crate::{ColumnRange, RowRange, TableBand, TableSpec};

pub const MAX_TOKENS: usize = 10_000;
pub const MAX_FORMULA_BYTES: usize = 32_768;
pub const MAX_FORMULA_CHARS: usize = 8192;
const TABLE_ESCAPE_CHARS: &[char] = &['[', ']', '#', '@', '\''];
const TABLE_BRACKET_CHARS: &[char] = &[
    '\t', '\n', '\r', ' ', ',', ':', '.', '[', ']', '#', '@', '\'', '"', '{', '}', '$', '^', '&',
    '*', '+', '=', '-', '>', '<', '/',
];
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
    /// `Table[...]`: a structured reference to a table's rows and columns.
    TableRef {
        sheet: Option<String>,
        table: String,
        spec: TableSpec,
        selector_start: usize,
    },
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
    if input.len() > MAX_FORMULA_BYTES {
        return Err(ParseError::new(
            MAX_FORMULA_BYTES,
            "formula length exceeded cap",
        ));
    }
    Lexer {
        input,
        pos: 0,
        external_prefixes: false,
        local_table_refs: false,
    }
    .run()
}

/// [`lex`] for rewriting formula text rather than evaluating it. It also
/// reads the prefix excel stores before a reference into another workbook,
/// `[1]` before `Sheet!A1` and `[1]!` before a name, as an
/// [`TokKind::Ident`] spanning just that prefix, so the text around such a
/// reference still tokenizes. The parser never sees these tokens.
pub fn lex_for_rewrite(input: &str) -> Result<Vec<Token>, ParseError> {
    if input.len() > MAX_FORMULA_BYTES {
        return Err(ParseError::new(
            MAX_FORMULA_BYTES,
            "formula length exceeded cap",
        ));
    }
    Lexer {
        input,
        pos: 0,
        external_prefixes: true,
        local_table_refs: false,
    }
    .run()
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
    let folded = input.to_lowercase();
    let column = column.to_lowercase();
    (local || folded.contains(&format!("{}[", table.to_lowercase())))
        && (folded.contains(&column)
            || folded.replace('\'', "").contains(&column.replace('\'', "")))
}

#[doc(hidden)]
pub fn rewrite_table_columns<'a>(
    input: &str,
    caption: impl Fn(Option<&str>, &str, &str) -> Option<&'a str>,
) -> Result<Option<String>, ParseError> {
    if input.len() > MAX_FORMULA_BYTES {
        return Err(ParseError::new(
            MAX_FORMULA_BYTES,
            "formula length exceeded cap",
        ));
    }
    let mut lexer = Lexer {
        input,
        pos: 0,
        external_prefixes: true,
        local_table_refs: true,
    };
    let tokens = lexer.run()?;
    let mut spans = Vec::new();
    for (index, token) in tokens.iter().enumerate() {
        let TokKind::TableRef {
            sheet,
            table,
            selector_start,
            ..
        } = &token.kind
        else {
            continue;
        };
        let external = (index > 0
            && matches!(&tokens[index - 1].kind, TokKind::Ident(word) if word.starts_with('[')))
            || sheet.as_deref().is_some_and(|sheet| sheet.starts_with('['));
        if external {
            continue;
        }
        lexer.pos = *selector_start;
        lexer.read_table_spec(token.start, |column, span| {
            if let Some(new) = caption(sheet.as_deref(), table, column) {
                let prefix = &input[*selector_start + 1..span.start];
                let needs_brackets = new.contains(TABLE_BRACKET_CHARS);
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
        })?;
    }
    if spans.is_empty() {
        return Ok(None);
    }
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

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
    external_prefixes: bool,
    local_table_refs: bool,
}

impl Lexer<'_> {
    fn run(&mut self) -> Result<Vec<Token>, ParseError> {
        let mut out = Vec::new();
        loop {
            self.skip_ws();
            let start = self.pos;
            let Some(c) = self.peek() else { break };
            let kind = match c {
                '0'..='9' if self.starts_row_range() => self.lex_word_or_ref()?,
                '0'..='9' | '.' => self.lex_number()?,
                '"' => self.lex_string()?,
                '#' => self.lex_error_literal()?,
                '\'' => self.lex_quoted_sheet_ref()?,
                '[' if self.workbook_function_end().is_some() => self.lex_workbook_function(),
                '[' if self.external_prefix_end().is_some() => self.lex_external_prefix(),
                '[' if self.local_table_refs => self.lex_table_ref(None, String::new(), start)?,
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
        while let Some(c) = self.peek() {
            if c == ' ' || c == '\t' || c == '\n' || c == '\r' {
                self.bump();
            } else {
                break;
            }
        }
    }

    /// read a maximal name/reference run.
    fn read_word(&mut self) -> String {
        let start = self.pos;
        while let Some(c) = self.peek() {
            if c.is_alphanumeric() || matches!(c, '_' | '.' | '$' | '\\') {
                self.bump();
            } else {
                break;
            }
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

    /// where `[1]!Foo` ends when one stands here before its `(`: the form
    /// excel stores a call to a function of another workbook, an add-in's, in.
    fn workbook_function_end(&self) -> Option<usize> {
        let rest = self.input[self.pos..].strip_prefix('[')?;
        let (index, name) = rest.split_once("]!")?;
        let word = name
            .find(|c: char| !c.is_alphanumeric() && !matches!(c, '_' | '.'))
            .unwrap_or(name.len());
        let named = !index.is_empty() && index.bytes().all(|b| b.is_ascii_digit()) && word > 0;
        (named && name[word..].starts_with('('))
            .then_some(self.pos + "[".len() + index.len() + "]!".len() + word)
    }

    /// where `[1]` or `[1]!` ends when one stands here and this lexer reads
    /// the prefixes of references into other workbooks.
    fn external_prefix_end(&self) -> Option<usize> {
        if !self.external_prefixes {
            return None;
        }
        let rest = self.input[self.pos..].strip_prefix('[')?;
        let (index, after) = rest.split_once(']')?;
        if index.is_empty() || !index.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        let bang = usize::from(after.starts_with('!'));
        if self.local_table_refs
            && bang == 0
            && !after.starts_with(|c: char| c.is_alphabetic() || matches!(c, '\'' | '_' | '\\'))
        {
            return None;
        }
        Some(self.pos + "[".len() + index.len() + "]".len() + bang)
    }

    fn lex_external_prefix(&mut self) -> TokKind {
        let start = self.pos;
        self.pos = self.external_prefix_end().unwrap_or(self.pos);
        TokKind::Ident(self.input[start..self.pos].to_string())
    }

    fn lex_workbook_function(&mut self) -> TokKind {
        let start = self.pos;
        self.pos = self.workbook_function_end().unwrap_or(self.pos);
        TokKind::Ident(self.input[start..self.pos].to_string())
    }

    /// `'Quoted Sheet'!A1` or `'Quoted Sheet'!A1:B2`. `''` escapes a quote.
    fn lex_quoted_sheet_ref(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        self.bump();
        let mut name = String::new();
        loop {
            match self.bump() {
                None => return Err(ParseError::new(start, "unterminated sheet name")),
                Some('\'') => {
                    if self.peek() == Some('\'') {
                        self.bump();
                        name.push('\'');
                    } else {
                        break;
                    }
                }
                Some(c) => name.push(c),
            }
        }
        if self.peek() != Some('!') {
            return Err(ParseError::new(
                self.pos,
                "expected '!' after quoted sheet name",
            ));
        }
        self.bump();
        self.lex_reference(Some(name), start)
    }

    /// classify a word starting with a letter or `$`: sheet-qualified ref,
    /// cell ref, range, boolean literal, or a bare name/function.
    fn lex_word_or_ref(&mut self) -> Result<TokKind, ParseError> {
        let start = self.pos;
        let word = self.read_word();

        if self.peek() == Some('!') {
            self.bump();
            return self.lex_reference(Some(word), start);
        }

        if self.peek() == Some('[') {
            return self.lex_table_ref(None, word, start);
        }

        // a name followed by `(` is a function even when ref-shaped (e.g. LOG10)
        // or a literal spelling (`TRUE()`)
        if self.peek() == Some('(') {
            return Ok(TokKind::Ident(word));
        }

        if word.eq_ignore_ascii_case("TRUE") {
            return Ok(TokKind::Bool(true));
        }
        if word.eq_ignore_ascii_case("FALSE") {
            return Ok(TokKind::Bool(false));
        }

        if let Some(kind) = self.try_range_token(&None, &word) {
            return Ok(kind);
        }

        match CellRef::parse_a1(&word.to_ascii_uppercase()) {
            Ok(cell) => Ok(TokKind::Ref { sheet: None, cell }),
            Err(_) => Ok(TokKind::Ident(word)),
        }
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
        let spec = self.read_table_spec(start, |_, _| {})?;
        Ok(TokKind::TableRef {
            sheet,
            table,
            spec,
            selector_start,
        })
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

    /// parse the reference part after a resolved sheet qualifier.
    fn lex_reference(
        &mut self,
        sheet: Option<String>,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        if self.input[self.pos..].starts_with("#REF!") {
            return self.lex_error_literal();
        }
        let word = self.read_word();
        if word.is_empty() {
            return Err(ParseError::new(start, "expected reference or defined name"));
        }
        if self.local_table_refs && self.peek() == Some('[') {
            return self.lex_table_ref(sheet, word, start);
        }
        if let Some(kind) = self.try_range_token(&sheet, &word) {
            return Ok(kind);
        }
        match CellRef::parse_a1(&word.to_ascii_uppercase()) {
            Ok(cell) => Ok(TokKind::Ref { sheet, cell }),
            Err(_) => Ok(TokKind::Name {
                scope: sheet,
                name: word,
            }),
        }
    }

    /// `word:end` as one token when `end` is also written out as a reference.
    /// anything else leaves the `:` for the parser, which reads it as the
    /// range operator — `A1:INDEX(..)` joins two references at evaluation
    /// time, not here.
    fn try_range_token(&mut self, sheet: &Option<String>, word: &str) -> Option<TokKind> {
        if !self.input[self.pos..].trim_start().starts_with(':') {
            return None;
        }
        let save = self.pos;
        if column(word).is_ok() {
            self.skip_ws();
            self.bump();
            self.skip_ws();
            let end = self.read_word();
            match ColumnRange::parse_a1(&format!("{word}:{end}")) {
                Ok(range) => {
                    return Some(TokKind::ColumnRange {
                        sheet: sheet.clone(),
                        range,
                    });
                }
                Err(_) => self.pos = save,
            }
        }
        if row(word).is_ok() {
            self.skip_ws();
            self.bump();
            self.skip_ws();
            let end = self.read_word();
            match RowRange::parse_a1(&format!("{word}:{end}")) {
                Ok(range) => {
                    return Some(TokKind::RowRange {
                        sheet: sheet.clone(),
                        range,
                    });
                }
                Err(_) => self.pos = save,
            }
        }
        let start = CellRef::parse_a1(&word.to_ascii_uppercase()).ok()?;
        if self.peek() != Some(':') {
            return None;
        }
        self.bump();
        let end = self.read_word();
        match CellRef::parse_a1(&end.to_ascii_uppercase()) {
            Ok(end) => Some(TokKind::Range {
                sheet: sheet.clone(),
                range: CellRange::new(start, end),
            }),
            Err(_) => {
                self.pos = save;
                None
            }
        }
    }

    /// whether the digits at the cursor open a `12:34` whole-row reference
    /// rather than a plain number.
    fn starts_row_range(&self) -> bool {
        let rest = &self.input[self.pos..];
        let after_digits = rest.trim_start_matches(|c: char| c.is_ascii_digit());
        let Some(tail) = after_digits.trim_start().strip_prefix(':') else {
            return false;
        };
        let tail = tail.trim_start();
        let tail = tail.strip_prefix('$').unwrap_or(tail);
        tail.starts_with(|c: char| c.is_ascii_digit())
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

    /// Only the rewrite lexer reads the prefix of a reference into another
    /// workbook, as one token spanning `[1]` or `[1]!`; the text around it
    /// tokenizes as usual, and the evaluating lexer still refuses it.
    #[test]
    fn the_rewrite_lexer_reads_external_workbook_prefixes() {
        for (source, prefix, rest) in [
            ("[1]Data!$A$1+Sales[Amt]", "[1]", 3),
            ("Sales[[#This Row],[Amt]]*[12]Rates!B2", "[12]", 1),
            ("[3]!Total+1", "[3]!", 3),
        ] {
            assert!(lex(source).is_err(), "{source}");
            let tokens = lex_for_rewrite(source).unwrap();
            let at = tokens
                .iter()
                .position(|t| t.kind == TokKind::Ident(prefix.into()))
                .unwrap();
            assert_eq!(&source[tokens[at].start..tokens[at].end], prefix);
            assert_eq!(tokens.len() - at - 1, rest, "{source}");
        }
        assert!(matches!(
            lex_for_rewrite("Sales[Amt]").unwrap()[0].kind,
            TokKind::TableRef { .. }
        ));
        assert!(lex_for_rewrite("[x]Data!A1").is_err());
        assert!(lex_for_rewrite("[]Data!A1").is_err());
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
