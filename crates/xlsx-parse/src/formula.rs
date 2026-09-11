//! Expand shared formulas while preserving their authored text and anchors.

use std::collections::BTreeMap;

use quick_xml::events::BytesStart;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS, col_to_letters};
use xlsx_model::{CellRange, CellRef, Sheet};

use crate::ParseError;
use crate::xml::attr;

const MAX_FORMULA_BYTES: usize = 32_768;
const MAX_SHARED_FORMULA_BYTES: usize = 32 * 1024 * 1024;

struct Master {
    origin: CellRef,
    range: CellRange,
    source: String,
}

pub(crate) struct SharedFormulas {
    masters: BTreeMap<u32, Master>,
    members: BTreeMap<(u32, u32), u32>,
    remaining: usize,
}

impl Default for SharedFormulas {
    fn default() -> Self {
        Self {
            masters: BTreeMap::new(),
            members: BTreeMap::new(),
            remaining: MAX_SHARED_FORMULA_BYTES,
        }
    }
}

impl SharedFormulas {
    pub(crate) fn record(
        &mut self,
        element: &BytesStart<'_>,
        origin: CellRef,
        source: &str,
    ) -> Result<(), ParseError> {
        let key = (origin.row, origin.col);
        if attr(element, b"t")?.as_deref() != Some("shared") {
            self.members.remove(&key);
            return Ok(());
        }
        let index = attr(element, b"si")?
            .and_then(|value| value.trim().parse::<u32>().ok())
            .ok_or_else(|| malformed("shared formula has no valid group index"))?;
        self.members.insert(key, index);
        if let Some(reference) = attr(element, b"ref")? {
            let range = CellRange::parse_a1(&reference)
                .map_err(|_| malformed("shared formula has an invalid range"))?;
            if !range.contains(origin) || source.trim().is_empty() {
                return Err(malformed("shared formula has an invalid master"));
            }
            if self.masters.contains_key(&index) {
                return Err(malformed("shared formula group has multiple masters"));
            }
            if source.len() > MAX_FORMULA_BYTES {
                return Err(malformed("shared formula exceeds the formula length limit"));
            }
            self.remaining = self
                .remaining
                .checked_sub(source.len())
                .ok_or_else(budget_error)?;
            self.masters.insert(
                index,
                Master {
                    origin,
                    range,
                    source: source.to_owned(),
                },
            );
        }
        Ok(())
    }

    pub(crate) fn resolve(mut self, sheet: &mut Sheet) -> Result<(), ParseError> {
        for ((row, col), index) in self.members {
            let at = CellRef::new(row, col);
            let master = self
                .masters
                .get(&index)
                .ok_or_else(|| malformed("shared formula group has no master"))?;
            if !master.range.contains(at) {
                return Err(malformed("shared formula is outside its master's range"));
            }
            if (row, col) == (master.origin.row, master.origin.col) {
                continue;
            }
            let formula = translate(
                &master.source,
                i64::from(row) - i64::from(master.origin.row),
                i64::from(col) - i64::from(master.origin.col),
                self.remaining.min(MAX_FORMULA_BYTES),
            )?;
            self.remaining -= formula.len();
            let cell = sheet
                .cell_mut(at)
                .ok_or_else(|| malformed("shared formula cell is missing"))?;
            cell.formula = Some(formula);
        }
        Ok(())
    }
}

fn malformed(message: &str) -> ParseError {
    ParseError::Malformed(message.to_owned())
}

fn budget_error() -> ParseError {
    malformed("shared formula expansion exceeds the byte limit")
}

fn append(output: &mut String, text: &str, limit: usize) -> Result<(), ParseError> {
    if text.len() > limit.saturating_sub(output.len()) {
        return Err(budget_error());
    }
    output.push_str(text);
    Ok(())
}

fn translate(source: &str, rows: i64, cols: i64, limit: usize) -> Result<String, ParseError> {
    if source.len() > limit {
        return Err(budget_error());
    }
    let mut output = String::with_capacity(source.len());
    let mut index = 0;
    while index < source.len() {
        let start = index;
        let character = source[index..].chars().next().expect("within source");
        if character == '"' {
            index = skip_quoted(source, index, b'"');
            append(&mut output, &source[start..index], limit)?;
        } else if separator(character) {
            index += character.len_utf8();
            append(&mut output, &source[start..index], limit)?;
        } else {
            index = operand_end(source, index);
            let operand = &source[start..index];
            let function = source[index..].trim_start().starts_with('(');
            translate_operand(&mut output, operand, rows, cols, limit, function)?;
        }
    }
    Ok(output)
}

fn separator(character: char) -> bool {
    character.is_whitespace()
        || matches!(
            character,
            '+' | '-'
                | '*'
                | '/'
                | '^'
                | '&'
                | '%'
                | '='
                | '<'
                | '>'
                | '('
                | ')'
                | ','
                | ';'
                | '{'
                | '}'
                | '@'
                | '#'
        )
}

fn skip_quoted(source: &str, start: usize, quote: u8) -> usize {
    let bytes = source.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == quote {
            index += 1;
            if bytes.get(index) != Some(&quote) {
                return index;
            }
        }
        index += 1;
    }
    bytes.len()
}

fn skip_brackets(source: &str, start: usize) -> usize {
    let bytes = source.as_bytes();
    let mut depth = 1;
    let mut index = start + 1;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' => index += usize::from(index + 1 < bytes.len()),
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return index + 1;
                }
            }
            _ => {}
        }
        index += 1;
    }
    bytes.len()
}

fn operand_end(source: &str, mut index: usize) -> usize {
    while index < source.len() {
        let character = source[index..].chars().next().expect("within source");
        match character {
            '\'' => index = skip_quoted(source, index, b'\''),
            '[' => index = skip_brackets(source, index),
            '"' => break,
            c if c.is_whitespace() => {
                let following = source[index..].trim_start();
                if !source[..index].trim_end().ends_with(':') && !following.starts_with(':') {
                    break;
                }
                index = source.len() - following.len();
            }
            c if separator(c) => break,
            c => index += c.len_utf8(),
        }
    }
    index
}

fn punctuation(source: &str, punctuation: u8) -> Vec<usize> {
    let bytes = source.as_bytes();
    let mut positions = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' | b'"' => index = skip_quoted(source, index, bytes[index]),
            b'[' => index = skip_brackets(source, index),
            byte => {
                if byte == punctuation {
                    positions.push(index);
                }
                index += 1;
            }
        }
    }
    positions
}

fn address_part(source: &str) -> (&str, &str) {
    match punctuation(source, b'!').last() {
        Some(index) => source.split_at(index + 1),
        None => ("", source),
    }
}

fn column(source: &str) -> Option<CellRef> {
    let letters = source.strip_prefix('$').unwrap_or(source);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    CellRef::parse_a1(&format!("{}1", source.to_ascii_uppercase())).ok()
}

fn row(source: &str) -> Option<CellRef> {
    let digits = source.strip_prefix('$').unwrap_or(source);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    CellRef::parse_a1(&format!("A{source}")).ok()
}

fn shifted(index: u32, absolute: bool, delta: i64, limit: u32) -> Option<u32> {
    let moved = i64::from(index).checked_add(if absolute { 0 } else { delta })?;
    (0..i64::from(limit))
        .contains(&moved)
        .then_some(moved as u32)
}

fn translate_operand(
    output: &mut String,
    operand: &str,
    rows: i64,
    cols: i64,
    limit: usize,
    function: bool,
) -> Result<(), ParseError> {
    let mut parts = Vec::new();
    let mut start = 0;
    for colon in punctuation(operand, b':') {
        parts.push(&operand[start..colon]);
        start = colon + 1;
    }
    parts.push(&operand[start..]);
    let three_dimensional = parts.len() > 1
        && address_part(parts[0]).0.is_empty()
        && !address_part(parts[1]).0.is_empty()
        && CellRef::parse_a1(&parts[0].trim().to_ascii_uppercase()).is_err();
    let addresses = &parts[usize::from(three_dimensional)..];
    let whole_columns = addresses.len() > 1
        && addresses
            .iter()
            .all(|part| column(address_part(part).1.trim()).is_some());
    let whole_rows = addresses.len() > 1
        && addresses
            .iter()
            .all(|part| row(address_part(part).1.trim()).is_some());
    for (index, part) in parts.iter().enumerate() {
        if index > 0 {
            append(output, ":", limit)?;
        }
        if (index == 0 && three_dimensional) || (function && index + 1 == parts.len()) {
            append(output, part, limit)?;
            continue;
        }
        let (prefix, address) = address_part(part);
        let trimmed = address.trim();
        let parsed = if whole_columns {
            column(trimmed)
        } else if whole_rows {
            row(trimmed)
        } else {
            CellRef::parse_a1(&trimmed.to_ascii_uppercase()).ok()
        };
        let Some(cell) = parsed else {
            append(output, part, limit)?;
            continue;
        };
        let moved_row = shifted(cell.row, cell.abs_row || whole_columns, rows, MAX_ROWS);
        let moved_col = shifted(cell.col, cell.abs_col || whole_rows, cols, MAX_COLS);
        let replacement = match (moved_row, moved_col) {
            (Some(row), Some(col)) if row == cell.row && col == cell.col => trimmed.to_owned(),
            (Some(_), Some(col)) if whole_columns => {
                format!(
                    "{}{}",
                    if cell.abs_col { "$" } else { "" },
                    col_to_letters(col)
                )
            }
            (Some(row), Some(_)) if whole_rows => {
                format!("{}{}", if cell.abs_row { "$" } else { "" }, row + 1)
            }
            (Some(row), Some(col)) => CellRef { row, col, ..cell }.to_a1(),
            _ => "#REF!".to_owned(),
        };
        append(output, prefix, limit)?;
        append(
            output,
            &address[..address.len() - address.trim_start().len()],
            limit,
        )?;
        append(output, &replacement, limit)?;
        append(output, &address[address.trim_end().len()..], limit)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifts_references_without_rewriting_other_formula_text() {
        for (source, expected) in [
            (r#"$B$6&"A"&$B$4&"B"&$D$2"#, r#"$B$6&"A"&$B$4&"B"&$D$2"#),
            ("A1+$B2+C$3+$D$4", "D3+$B4+F$3+$D$4"),
            ("SUM(A1:B2,$C$3:D$4)", "SUM(D3:E4,$C$3:G$4)"),
            ("SUM(B2:A1)", "SUM(E4:D3)"),
            ("Sheet1!A1+'Other Sheet'!$B2", "Sheet1!D3+'Other Sheet'!$B4"),
            ("'A1''s ! sheet'!C3", "'A1''s ! sheet'!F5"),
            ("'A1:B2'!C3:D4", "'A1:B2'!F5:G6"),
            ("Sheet1:Sheet3!A1:B2", "Sheet1:Sheet3!D3:E4"),
            ("Sheet1!A1:Sheet2!B2", "Sheet1!D3:Sheet2!E4"),
            ("A1:Sheet2!B2", "D3:Sheet2!E4"),
            ("[1]Sheet1!A1+[1]Sheet1!$B$2", "[1]Sheet1!D3+[1]Sheet1!$B$2"),
            (
                "'C:\\A1\\[Book.xlsx]Sheet1'!A1",
                "'C:\\A1\\[Book.xlsx]Sheet1'!D3",
            ),
            (r#"IF(A1="B2","A1""C3",D4)"#, r#"IF(D3="B2","A1""C3",G6)"#),
            (
                "LOG10(A1)+LOG10 (B2)+1E3+2.5e-4",
                "LOG10(D3)+LOG10 (E4)+1E3+2.5e-4",
            ),
            (
                "_xlfn.FOO(A1,Rate_A1,A1_rate,éA1,A1.Thing)",
                "_xlfn.FOO(D3,Rate_A1,A1_rate,éA1,A1.Thing)",
            ),
            (
                "Table1[A1]+Table1[[#Headers],[B2]]+A1",
                "Table1[A1]+Table1[[#Headers],[B2]]+D3",
            ),
            ("Table1[A1']B2]+C3", "Table1[A1']B2]+F5"),
            ("@A1+A1#+#REF!+#N/A", "@D3+D3#+#REF!+#N/A"),
            (r#"SUM({1,2;3,4})+"A1"&A1"#, r#"SUM({1,2;3,4})+"A1"&D3"#),
            ("SUM(A:C,$D:$E,1:3,$4:$5)", "SUM(D:F,$D:$E,3:5,$4:$5)"),
            ("SUM(A : $C, 1 : $3)", "SUM(D : $C, 3 : $3)"),
            ("SUM(Sheet1!A:Sheet2!B)", "SUM(Sheet1!D:Sheet2!E)"),
            ("SUM(Sheet1:Sheet3!A:C)", "SUM(Sheet1:Sheet3!D:F)"),
            ("SUM(A:A B:B)", "SUM(D:D E:E)"),
            ("SUM(A1 : B$2)", "SUM(D3 : E$2)"),
            ("SUM(A1:INDEX(B1:B9,2))", "SUM(D3:INDEX(E3:E11,2))"),
            ("SUM(INDEX(B1:B9,2):A9)", "SUM(INDEX(E3:E11,2):D11)"),
            ("SUM(A1:OFFSET(B1,1,0))", "SUM(D3:OFFSET(E3,1,0))"),
            ("a1+$b$2", "D3+$b$2"),
        ] {
            assert_eq!(
                translate(source, 2, 3, MAX_FORMULA_BYTES).unwrap(),
                expected,
                "{source}"
            );
        }
    }

    #[test]
    fn shifts_backwards_and_reports_out_of_grid_references() {
        assert_eq!(
            translate("A1+$A$1+B$1+$A2+B2", -1, -1, MAX_FORMULA_BYTES).unwrap(),
            "#REF!+$A$1+A$1+$A1+A1"
        );
        assert_eq!(
            translate("XFD1048576+$XFD$1048576", 1, 1, MAX_FORMULA_BYTES).unwrap(),
            "#REF!+$XFD$1048576"
        );
        assert_eq!(
            translate("SUM(A:B,1:2)", -1, -1, MAX_FORMULA_BYTES).unwrap(),
            "SUM(#REF!:A,#REF!:1)"
        );
    }

    #[test]
    fn refuses_expansion_past_the_remaining_budget() {
        assert!(translate("A9", 1, 0, 2).is_err());
        assert_eq!(translate("A9", 1, 0, 3).unwrap(), "A10");
        assert!(translate("$A$1", 0, 0, 3).is_err());
    }

    #[test]
    fn charges_all_shared_groups_against_one_sheet_budget() {
        let mut formulas = SharedFormulas {
            remaining: 3,
            ..SharedFormulas::default()
        };
        let master =
            BytesStart::new("f").with_attributes([("t", "shared"), ("si", "0"), ("ref", "A1:A2")]);
        formulas.record(&master, CellRef::new(0, 0), "B1").unwrap();
        let follower = BytesStart::new("f").with_attributes([("t", "shared"), ("si", "0")]);
        formulas.record(&follower, CellRef::new(1, 0), "").unwrap();
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(
            CellRef::new(1, 0),
            xlsx_model::Cell {
                formula: Some(String::new()),
                ..Default::default()
            },
        );
        assert!(formulas.resolve(&mut sheet).is_err());

        let mut formulas = SharedFormulas {
            remaining: 3,
            ..SharedFormulas::default()
        };
        formulas.record(&master, CellRef::new(0, 0), "B1").unwrap();
        let second =
            BytesStart::new("f").with_attributes([("t", "shared"), ("si", "1"), ("ref", "C1:C2")]);
        assert!(formulas.record(&second, CellRef::new(0, 2), "D1").is_err());
    }
}
