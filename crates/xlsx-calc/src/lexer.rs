//! formula lexer: turns source (without the leading `=`) into position-tagged
//! tokens, resolving reference/range/sheet disambiguation.

use std::fmt;

use xlsx_model::{CellRange, CellRef, ErrorValue};

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
    Ref {
        sheet: Option<String>,
        cell: CellRef,
    },
    Range {
        sheet: Option<String>,
        range: CellRange,
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
}

/// tokenize a formula source string.
pub fn lex(input: &str) -> Result<Vec<Token>, ParseError> {
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
                '0'..='9' | '.' => self.lex_number()?,
                '"' => self.lex_string()?,
                '#' => self.lex_error_literal()?,
                '\'' => self.lex_quoted_sheet_ref()?,
                '(' => self.punct(TokKind::LParen),
                ')' => self.punct(TokKind::RParen),
                ',' => self.punct(TokKind::Comma),
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
                c if c.is_ascii_alphabetic() || c == '$' => self.lex_word_or_ref()?,
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

    /// read a maximal `[A-Za-z0-9_.$]` run (ref parts, sheet names, func names).
    fn read_word(&mut self) -> String {
        let start = self.pos;
        while let Some(c) = self.peek() {
            if c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$' {
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
        text.parse::<f64>()
            .map(TokKind::Num)
            .map_err(|_| ParseError::new(start, format!("malformed number {text:?}")))
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

        if self.peek() == Some(':') && CellRef::parse_a1(&word).is_ok() {
            return self.finish_range(None, &word, start);
        }

        match CellRef::parse_a1(&word) {
            Ok(cell) => Ok(TokKind::Ref { sheet: None, cell }),
            Err(_) => Ok(TokKind::Ident(word)),
        }
    }

    /// parse the reference part after a resolved sheet qualifier.
    fn lex_reference(
        &mut self,
        sheet: Option<String>,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        let word = self.read_word();
        if self.peek() == Some(':') {
            return self.finish_range(sheet, &word, start);
        }
        let cell = CellRef::parse_a1(&word)
            .map_err(|e| ParseError::new(start, format!("invalid reference {word:?}: {e}")))?;
        Ok(TokKind::Ref { sheet, cell })
    }

    /// consume `:end` and build a range from an already-read start segment.
    fn finish_range(
        &mut self,
        sheet: Option<String>,
        start_word: &str,
        start: usize,
    ) -> Result<TokKind, ParseError> {
        self.bump();
        let end_word = self.read_word();
        let a = CellRef::parse_a1(start_word).map_err(|e| {
            ParseError::new(start, format!("invalid range start {start_word:?}: {e}"))
        })?;
        let b = CellRef::parse_a1(&end_word)
            .map_err(|e| ParseError::new(start, format!("invalid range end {end_word:?}: {e}")))?;
        Ok(TokKind::Range {
            sheet,
            range: CellRange::new(a, b),
        })
    }
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
    fn rejects_unterminated_string() {
        let err = lex("\"abc").unwrap_err();
        assert_eq!(err.pos, 0);
    }

    #[test]
    fn rejects_unexpected_char() {
        assert!(lex("1 ~ 2").is_err());
    }
}
