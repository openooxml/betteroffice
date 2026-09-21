//! formula lexer: turns source (without the leading `=`) into position-tagged
//! tokens, resolving reference/range/sheet disambiguation.

use std::fmt;

use xlsx_model::{CellRange, CellRef, ErrorValue};

use crate::reference::{MAX_TABLE_SPEC_ITEMS, column, row};
use crate::{ColumnRange, RowRange, TableBand, TableSpec};

pub const MAX_TOKENS: usize = 10_000;
pub const MAX_FORMULA_BYTES: usize = 32_768;

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
        table: String,
        spec: TableSpec,
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
    Lexer { input, pos: 0 }.run()
}

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
}

impl Lexer<'_> {
    fn run(mut self) -> Result<Vec<Token>, ParseError> {
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
        let value = text
            .parse::<f64>()
            .map_err(|_| ParseError::new(start, format!("malformed number {text:?}")))?;
        if !value.is_finite() {
            return Err(ParseError::new(start, "number is outside the finite range"));
        }
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
            return self.lex_table_ref(word, start);
        }

        if word.eq_ignore_ascii_case("TRUE") {
            return Ok(TokKind::Bool(true));
        }
        if word.eq_ignore_ascii_case("FALSE") {
            return Ok(TokKind::Bool(false));
        }

        // a name followed by `(` is a function even when ref-shaped (e.g. LOG10)
        if self.peek() == Some('(') {
            return Ok(TokKind::Ident(word));
        }

        if let Some(kind) = self.try_range_token(&None, &word) {
            return Ok(kind);
        }

        match CellRef::parse_a1(&word) {
            Ok(cell) => Ok(TokKind::Ref { sheet: None, cell }),
            Err(_) => Ok(TokKind::Ident(word)),
        }
    }

    /// `Table[...]`: either one bare item (`Table[Col]`, `Table[#Headers]`,
    /// `Table[@Col]`) or a bracketed list (`Table[[#This Row],[Col]]`).
    fn lex_table_ref(&mut self, table: String, start: usize) -> Result<TokKind, ParseError> {
        self.bump();
        let mut spec = TableSpec::default();
        self.skip_ws();
        if self.peek() == Some(']') {
            self.bump();
            return Ok(TokKind::TableRef { table, spec });
        }
        if self.peek() != Some('[') {
            let item = self.read_table_item(start)?;
            add_table_item(&mut spec, &item, start)?;
            return Ok(TokKind::TableRef { table, spec });
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
                    let item = self.read_table_item(start)?;
                    add_table_item(&mut spec, &item, start)?;
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
        Ok(TokKind::TableRef { table, spec })
    }

    /// read up to and including the item's closing `]`; `'` escapes the next
    /// character, which is how excel writes `[`, `]`, `#`, `@` and `'` in a
    /// column name.
    fn read_table_item(&mut self, start: usize) -> Result<String, ParseError> {
        let mut out = String::new();
        loop {
            match self.bump() {
                None => return Err(ParseError::new(start, "unterminated table reference")),
                Some(']') => break,
                Some('\'') => match self.bump() {
                    None => return Err(ParseError::new(start, "unterminated table reference")),
                    Some(escaped) => out.push(escaped),
                },
                Some(c) => out.push(c),
            }
        }
        Ok(out)
    }

    /// parse the reference part after a resolved sheet qualifier.
    fn lex_reference(
        &mut self,
        sheet: Option<String>,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        let word = self.read_word();
        if word.is_empty() {
            return Err(ParseError::new(start, "expected reference or defined name"));
        }
        if let Some(kind) = self.try_range_token(&sheet, &word) {
            return Ok(kind);
        }
        match CellRef::parse_a1(&word) {
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
        let start = CellRef::parse_a1(word).ok()?;
        if self.peek() != Some(':') {
            return None;
        }
        self.bump();
        let end = self.read_word();
        match CellRef::parse_a1(&end) {
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

    fn finish_rows(
        &mut self,
        sheet: Option<String>,
        start_word: &str,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        self.skip_ws();
        self.bump();
        self.skip_ws();
        let end_word = self.read_word();
        let range = RowRange::parse_a1(&format!("{start_word}:{end_word}"))
            .map_err(|error| ParseError::new(start, format!("invalid row range: {error}")))?;
        Ok(TokKind::RowRange { sheet, range })
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

/// classify one bracket item as a band keyword, an `@` shorthand, or a column
/// name, and fold it into `spec`.
fn add_table_item(spec: &mut TableSpec, item: &str, start: usize) -> Result<(), ParseError> {
    let mut item = item.trim();
    if item.is_empty() {
        return Ok(());
    }
    if let Some(rest) = item.strip_prefix('@') {
        spec.bands.push(TableBand::ThisRow);
        item = rest.trim();
        if item.is_empty() {
            return Ok(());
        }
    }
    if item.starts_with('#') {
        let band = TableBand::parse(item)
            .ok_or_else(|| ParseError::new(start, "unknown table band selector"))?;
        spec.bands.push(band);
        return Ok(());
    }
    if spec.first_column.is_none() {
        spec.first_column = Some(item.to_string());
    } else if spec.last_column.is_none() {
        spec.last_column = Some(item.to_string());
    } else {
        return Err(ParseError::new(start, "too many table columns"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(src: &str) -> Vec<TokKind> {
        lex(src).unwrap().into_iter().map(|t| t.kind).collect()
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
        assert!(lex("1e999").unwrap_err().message.contains("finite"));
    }
}
