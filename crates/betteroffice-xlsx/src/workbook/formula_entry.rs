//! New formula text with a character XML 1.0 cannot carry, not even as a
//! character reference, would make a file Excel refuses to open. Inside a
//! text literal such a character is written as `CHAR(n)` instead, which keeps
//! the formula's value; anywhere else the text is refused.

use xlsx_calc::lexer::{TokKind, lex};

use crate::{Error, Result};

/// Excel's limit on a formula's length, in characters.
const MAX_FORMULA_CHARS: usize = 8192;

fn storable(ch: char) -> bool {
    matches!(ch, '\t' | '\n' | '\r') || (ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}')
}

fn refused(message: String) -> Error {
    Error::InvalidOperation(message)
}

/// `formula` with every text literal that holds a control character written
/// as a parenthesized concatenation, `"a␁b"` as `("a"&CHAR(1)&"b")`, or
/// `formula` itself when every character is one a file can store. Tab, line
/// feed and carriage return stay in the literal, since a file stores them as
/// character references.
pub(crate) fn storable_formula(formula: String) -> Result<String> {
    let Some(first) = formula.chars().find(|ch| !storable(*ch)) else {
        return Ok(formula);
    };
    let tokens = lex(&formula).map_err(|error| {
        refused(format!(
            "formula contains U+{:04X} and cannot be read ({error}), so it cannot be stored",
            first as u32
        ))
    })?;
    let mut out = String::with_capacity(formula.len() + 16);
    let mut copied = 0;
    let mut depth = 0usize;
    for token in &tokens {
        match &token.kind {
            TokKind::LBrace => depth += 1,
            TokKind::RBrace => depth = depth.saturating_sub(1),
            TokKind::Str(text) if text.chars().any(|ch| !storable(ch)) => {
                if depth > 0 {
                    let ch = text.chars().find(|ch| !storable(*ch)).unwrap_or(first);
                    return Err(refused(format!(
                        "formula contains U+{:04X} inside an array constant, which a workbook file cannot store",
                        ch as u32
                    )));
                }
                out.push_str(&formula[copied..token.start]);
                out.push_str(&concatenation(text)?);
                copied = token.end;
            }
            _ => {}
        }
    }
    out.push_str(&formula[copied..]);
    if let Some(ch) = out.chars().find(|ch| !storable(*ch)) {
        return Err(refused(format!(
            "formula contains U+{:04X} outside a text literal, which a workbook file cannot store",
            ch as u32
        )));
    }
    if out.chars().count() > MAX_FORMULA_CHARS {
        return Err(refused(format!(
            "formula with its control characters written as CHAR exceeds Excel's {MAX_FORMULA_CHARS}-character limit"
        )));
    }
    Ok(out)
}

/// The literal `text` as `("a"&CHAR(1)&"b")`, with no empty pieces.
fn concatenation(text: &str) -> Result<String> {
    let mut pieces = Vec::new();
    let mut run = String::new();
    for ch in text.chars() {
        if storable(ch) {
            run.push(ch);
            continue;
        }
        let code = ch as u32;
        if code == 0 || code > 0xff {
            return Err(refused(format!(
                "formula contains U+{code:04X}, which a workbook file cannot store and CHAR cannot produce"
            )));
        }
        if !run.is_empty() {
            pieces.push(format!(
                "\"{}\"",
                std::mem::take(&mut run).replace('"', "\"\"")
            ));
        }
        pieces.push(format!("CHAR({code})"));
    }
    if !run.is_empty() {
        pieces.push(format!("\"{}\"", run.replace('"', "\"\"")));
    }
    Ok(format!("({})", pieces.join("&")))
}

#[cfg(test)]
mod tests {
    use super::storable_formula;

    fn stored(formula: &str) -> String {
        storable_formula(formula.to_owned()).unwrap()
    }

    fn refusal(formula: &str) -> String {
        storable_formula(formula.to_owned())
            .unwrap_err()
            .to_string()
    }

    #[test]
    fn literals_with_control_characters_become_parenthesized_concatenations() {
        assert_eq!(stored("\"a\u{1}b\""), "(\"a\"&CHAR(1)&\"b\")");
        assert_eq!(stored("\"a\u{1}\u{2}b\""), "(\"a\"&CHAR(1)&CHAR(2)&\"b\")");
        assert_eq!(stored("\"\u{1}\""), "(CHAR(1))");
        assert_eq!(stored("\"\u{1}a\""), "(CHAR(1)&\"a\")");
        assert_eq!(
            stored("\"a\"\"\u{7}\"\"b\""),
            "(\"a\"\"\"&CHAR(7)&\"\"\"b\")"
        );
        assert_eq!(stored("-\"\u{1}\"&A1"), "-(CHAR(1))&A1");
        assert_eq!(
            stored("IF(A1=\"x\u{1f}\", \"y\" , 'My Sheet'!B2)"),
            "IF(A1=(\"x\"&CHAR(31)), \"y\" , 'My Sheet'!B2)"
        );
        assert_eq!(
            stored("CONCAT(\"a\u{1}\",Table1[Col],\"b\u{2}\")"),
            "CONCAT((\"a\"&CHAR(1)),Table1[Col],(\"b\"&CHAR(2)))"
        );
    }

    #[test]
    fn formulas_a_file_can_store_are_left_as_they_are() {
        for formula in [
            "\"a\rb\"",
            "\"a\nb\"&\"\tc\"",
            "A1+1",
            "{\"a\",\"b\"}",
            "\"é\u{2028}\"",
        ] {
            assert_eq!(stored(formula), formula);
        }
    }

    #[test]
    fn characters_with_no_storable_form_are_refused() {
        assert!(
            refusal("\"a\u{0}\"")
                .contains("U+0000, which a workbook file cannot store and CHAR cannot produce")
        );
        assert!(
            refusal("\"a\u{fffe}\"")
                .contains("U+FFFE, which a workbook file cannot store and CHAR cannot produce")
        );
        assert!(refusal("\"a\u{ffff}\"").contains("U+FFFF"));
        assert!(refusal("{\"a\u{1}\",1}").contains("U+0001 inside an array constant"));
        assert!(refusal("A1\u{1}").contains("U+0001 and cannot be read"));
        assert!(refusal("'Data\u{1}'!A1").contains("U+0001 outside a text literal"));
        let long = format!("\"{}\"", "\u{1}".repeat(1100));
        assert!(refusal(&long).contains("8192-character limit"));
    }
}
