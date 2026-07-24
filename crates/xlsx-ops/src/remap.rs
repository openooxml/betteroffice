//! rewriting stored formulas on row/column insert/delete: refs shift, ranges
//! clip, wholly deleted refs collapse to `#REF!`. runs before cells shift.

use std::collections::HashMap;

use xlsx_calc::lexer::MAX_FORMULA_BYTES;
use xlsx_calc::parse_formula;
use xlsx_calc::parser::Expr;
use xlsx_model::{CellRange, CellRef, ErrorValue, SheetId, Workbook};

use crate::apply::OpError;
use crate::apply::remap_ref;
use crate::op::{CellState, Op};

/// the outcome of remapping one reference or range under a structural op.
enum Remapped<T> {
    /// the op does not move this reference.
    Unchanged,
    /// the reference shifted (or a range clipped) to a new address.
    Moved(T),
    /// the reference (or the whole range) fell inside the deleted span.
    Deleted,
}

/// the sheet a structural op targets, or `None` for non-structural ops.
fn structural_target(op: &Op) -> Option<SheetId> {
    match *op {
        Op::InsertRows { sheet, .. }
        | Op::DeleteRows { sheet, .. }
        | Op::InsertCols { sheet, .. }
        | Op::DeleteCols { sheet, .. } => Some(sheet),
        _ => None,
    }
}

/// rewrite every workbook formula affected by `op`, in place, returning the
/// inverse `SetCell` ops that restore the rewritten formulas.
pub(crate) fn remap_formulas(wb: &mut Workbook, op: &Op) -> Result<Vec<Op>, OpError> {
    let Some(target) = structural_target(op) else {
        return Ok(Vec::new());
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(i, s)| (s.name.to_lowercase(), SheetId(i as u32)))
        .collect();

    let mut restores: Vec<Op> = Vec::new();
    let mut edits: Vec<(SheetId, CellRef, String)> = Vec::new();
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(i as u32);
        let matches =
            |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, target, &names);
        for (cell, c) in sheet.iter_cells() {
            let Some(src) = &c.formula else {
                continue;
            };
            let expr = parse_formula(src)
                .map_err(|_| OpError::FormulaNotRewritable { sheet: owner, cell })?;
            let mut changed = false;
            let new_expr = transform(&expr, op, &matches, &mut changed);
            if changed {
                let formula = new_expr.to_formula();
                parse_formula(&formula)
                    .map_err(|_| OpError::FormulaNotRewritable { sheet: owner, cell })?;
                restores.push(Op::SetCell {
                    sheet: owner,
                    at: cell,
                    cell: CellState::from(c),
                });
                edits.push((owner, cell, formula));
            }
        }
    }

    for (sheet, at, text) in edits {
        let s = wb.sheet_mut(sheet).expect("sheet exists during remap");
        if let Some(mut cell) = s.cell(at).cloned() {
            cell.formula = Some(text);
            s.set_cell(at, cell);
        }
    }
    Ok(restores)
}

pub(crate) fn remap_hyperlink_locations(wb: &mut Workbook, op: &Op) -> Vec<Op> {
    let Some(target) = structural_target(op) else {
        return Vec::new();
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| (sheet.name.to_lowercase(), SheetId(index as u32)))
        .collect();
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        let matches =
            |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, target, &names);
        let mut hyperlinks = sheet.hyperlinks.clone();
        let mut changed_sheet = false;
        for hyperlink in &mut hyperlinks {
            if hyperlink.external_target.is_some() {
                continue;
            }
            let Some(location) = &hyperlink.location else {
                continue;
            };
            let prefixed = location.starts_with('#');
            let source = location.strip_prefix('#').unwrap_or(location);
            let Ok(expr) = parse_formula(source) else {
                continue;
            };
            let mut changed = false;
            let rewritten = transform(&expr, op, &matches, &mut changed).to_formula();
            if changed {
                hyperlink.location = Some(if prefixed && !rewritten.starts_with('#') {
                    format!("#{rewritten}")
                } else {
                    rewritten
                });
                changed_sheet = true;
            }
        }
        if changed_sheet {
            restores.push(Op::SetHyperlinks {
                sheet: owner,
                hyperlinks: sheet.hyperlinks.clone(),
            });
            edits.push((owner, hyperlinks));
        }
    }
    for (sheet, hyperlinks) in edits {
        wb.sheet_mut(sheet)
            .expect("sheet exists during hyperlink remap")
            .hyperlinks = hyperlinks;
    }
    restores
}

pub(crate) fn remap_hyperlink_range(range: CellRange, op: &Op) -> Option<CellRange> {
    match remap_span(range, op) {
        Remapped::Unchanged => Some(range),
        Remapped::Moved(range) => Some(range),
        Remapped::Deleted => None,
    }
}

pub(crate) fn rename_sheet_references(
    wb: &mut Workbook,
    old_name: &str,
    new_name: &str,
) -> Result<Vec<(SheetId, CellRef, CellState)>, OpError> {
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        for (cell, stored) in sheet.iter_cells() {
            let Some(source) = &stored.formula else {
                continue;
            };
            let rewritten = rename_formula_sheet(source, old_name, new_name);
            if rewritten != *source {
                if rewritten.len() > MAX_FORMULA_BYTES {
                    return Err(OpError::FormulaNotRewritable { sheet: owner, cell });
                }
                restores.push((owner, cell, CellState::from(stored)));
                edits.push((owner, cell, rewritten));
            }
        }
    }
    for (sheet, cell, formula) in edits {
        let stored = wb
            .sheet_mut(sheet)
            .and_then(|sheet| sheet.cell(cell).cloned());
        if let Some(mut stored) = stored {
            stored.formula = Some(formula);
            wb.sheet_mut(sheet)
                .expect("sheet exists")
                .set_cell(cell, stored);
        }
    }
    Ok(restores)
}

pub(crate) fn rename_formula_sheet(source: &str, old_name: &str, new_name: &str) -> String {
    if old_name == new_name {
        return source.to_string();
    }
    let bytes = source.as_bytes();
    let mut replacements = Vec::new();
    let mut index = 0;
    let mut bracket_depth = 0_u32;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = skip_string(source, index);
            continue;
        }
        if bytes[index] == b'[' {
            bracket_depth = bracket_depth.saturating_add(1);
            index += 1;
            continue;
        }
        if bytes[index] == b']' {
            bracket_depth = bracket_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if bracket_depth != 0 {
            index += source[index..]
                .chars()
                .next()
                .map(char::len_utf8)
                .unwrap_or(1);
            continue;
        }
        let Some(first) = parse_sheet_token(source, index) else {
            index += source[index..]
                .chars()
                .next()
                .map(char::len_utf8)
                .unwrap_or(1);
            continue;
        };
        let external = first.start > 0 && bytes[first.start - 1] == b']';
        if bytes.get(first.end) == Some(&b'!') {
            if !external && sheet_names_equal(&first.name, old_name) {
                replacements.push((first.start, first.end));
            }
            index = first.end + 1;
            continue;
        }
        if bytes.get(first.end) == Some(&b':')
            && let Some(second) = parse_sheet_token(source, first.end + 1)
            && bytes.get(second.end) == Some(&b'!')
        {
            if !external && sheet_names_equal(&first.name, old_name) {
                replacements.push((first.start, first.end));
            }
            if !external && sheet_names_equal(&second.name, old_name) {
                replacements.push((second.start, second.end));
            }
            index = second.end + 1;
            continue;
        }
        index = first.end;
    }
    if replacements.is_empty() {
        return source.to_string();
    }
    let replacement = sheet_token(new_name);
    let mut output = String::with_capacity(source.len());
    let mut copied_until = 0;
    for (start, end) in replacements {
        output.push_str(&source[copied_until..start]);
        output.push_str(&replacement);
        copied_until = end;
    }
    output.push_str(&source[copied_until..]);
    output
}

struct ParsedSheetToken {
    start: usize,
    end: usize,
    name: String,
}

fn parse_sheet_token(source: &str, start: usize) -> Option<ParsedSheetToken> {
    let bytes = source.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut index = start + 1;
        let mut name = String::new();
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    name.push('\'');
                    index += 2;
                } else {
                    return Some(ParsedSheetToken {
                        start,
                        end: index + 1,
                        name,
                    });
                }
            } else {
                let character = source[index..].chars().next()?;
                name.push(character);
                index += character.len_utf8();
            }
        }
        return None;
    }
    let first = source[start..].chars().next()?;
    if !is_unquoted_sheet_char(first) {
        return None;
    }
    let mut end = start;
    while end < bytes.len() {
        let character = source[end..].chars().next()?;
        if !is_unquoted_sheet_char(character) {
            break;
        }
        end += character.len_utf8();
    }
    Some(ParsedSheetToken {
        start,
        end,
        name: source[start..end].to_string(),
    })
}

fn skip_string(source: &str, start: usize) -> usize {
    let bytes = source.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index += 1;
            if bytes.get(index) == Some(&b'"') {
                index += 1;
            } else {
                break;
            }
        } else {
            index += source[index..]
                .chars()
                .next()
                .map(char::len_utf8)
                .unwrap_or(1);
        }
    }
    index
}

fn sheet_names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

fn is_unquoted_sheet_char(character: char) -> bool {
    !character.is_whitespace()
        && !matches!(
            character,
            '"' | '\''
                | '!'
                | ':'
                | '+'
                | '-'
                | '*'
                | '/'
                | '^'
                | '&'
                | '='
                | '<'
                | '>'
                | '('
                | ')'
                | ','
                | '%'
                | '['
                | ']'
        )
}

fn sheet_token(name: &str) -> String {
    let simple = !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || character == '_')
        && !name
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
        && CellRef::parse_a1(name).is_err()
        && !is_r1c1_reference(name);
    if simple {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn is_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    if matches!(upper.as_str(), "R" | "C") {
        return true;
    }
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
    rest.as_bytes().get(digits) == Some(&b'C')
}

/// whether a reference read from `owner` points at the edited sheet `target`.
/// unqualified refs bind to `owner`; unknown sheet names never match.
fn resolves_to_target(
    ref_sheet: &Option<String>,
    owner: SheetId,
    target: SheetId,
    names: &HashMap<String, SheetId>,
) -> bool {
    let resolved = match ref_sheet {
        None => Some(owner),
        Some(name) => names.get(&name.to_lowercase()).copied(),
    };
    resolved == Some(target)
}

/// rebuild `expr`, remapping every reference to the edited sheet; sets
/// `changed` only when a reference actually moved.
fn transform(
    expr: &Expr,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    changed: &mut bool,
) -> Expr {
    match expr {
        Expr::Ref { sheet, cell } if matches_target(sheet) => match remap_cell(*cell, op) {
            Remapped::Unchanged => expr.clone(),
            Remapped::Moved(new_cell) => {
                *changed = true;
                Expr::Ref {
                    sheet: sheet.clone(),
                    cell: new_cell,
                }
            }
            Remapped::Deleted => {
                *changed = true;
                Expr::Error(ErrorValue::Ref)
            }
        },
        Expr::Range { sheet, range } if matches_target(sheet) => match remap_span(*range, op) {
            Remapped::Unchanged => expr.clone(),
            Remapped::Moved(new_range) => {
                *changed = true;
                Expr::Range {
                    sheet: sheet.clone(),
                    range: new_range,
                }
            }
            Remapped::Deleted => {
                *changed = true;
                Expr::Error(ErrorValue::Ref)
            }
        },
        Expr::Unary { op: u, expr: e } => Expr::Unary {
            op: *u,
            expr: Box::new(transform(e, op, matches_target, changed)),
        },
        Expr::Percent(e) => Expr::Percent(Box::new(transform(e, op, matches_target, changed))),
        Expr::Binary { op: b, lhs, rhs } => Expr::Binary {
            op: *b,
            lhs: Box::new(transform(lhs, op, matches_target, changed)),
            rhs: Box::new(transform(rhs, op, matches_target, changed)),
        },
        Expr::FuncCall { name, args } => Expr::FuncCall {
            name: name.clone(),
            args: args
                .iter()
                .map(|a| transform(a, op, matches_target, changed))
                .collect(),
        },
        _ => expr.clone(),
    }
}

/// remap a single-cell reference through the op.
fn remap_cell(cell: CellRef, op: &Op) -> Remapped<CellRef> {
    match remap_ref(cell, op) {
        Some(new_cell) if new_cell == cell => Remapped::Unchanged,
        Some(new_cell) => Remapped::Moved(new_cell),
        None => Remapped::Deleted,
    }
}

/// remap a range: inserts shift both corners; deletes clip the span, collapsing
/// to `#REF!` only when the whole span is deleted.
fn remap_span(range: CellRange, op: &Op) -> Remapped<CellRange> {
    match *op {
        Op::DeleteRows { at, count, .. } => clip_span(range, Axis::Row, at, count),
        Op::DeleteCols { at, count, .. } => clip_span(range, Axis::Col, at, count),
        // inserts only drop a corner on off-sheet-edge overflow
        _ => match (remap_ref(range.start, op), remap_ref(range.end, op)) {
            (Some(start), Some(end)) if start == range.start && end == range.end => {
                Remapped::Unchanged
            }
            (Some(start), Some(end)) => Remapped::Moved(CellRange { start, end }),
            _ => Remapped::Deleted,
        },
    }
}

enum Axis {
    Row,
    Col,
}

/// clip a range's span on one axis under a delete of `count` starting at `at`.
fn clip_span(range: CellRange, axis: Axis, at: u32, count: u32) -> Remapped<CellRange> {
    let (a, b) = match axis {
        Axis::Row => (range.start.row, range.end.row),
        Axis::Col => (range.start.col, range.end.col),
    };
    match clip_interval(a, b, at, count) {
        None => Remapped::Deleted,
        Some((na, nb)) if na == a && nb == b => Remapped::Unchanged,
        Some((na, nb)) => {
            let mut start = range.start;
            let mut end = range.end;
            match axis {
                Axis::Row => {
                    start.row = na;
                    end.row = nb;
                }
                Axis::Col => {
                    start.col = na;
                    end.col = nb;
                }
            }
            Remapped::Moved(CellRange { start, end })
        }
    }
}

/// clip the inclusive interval `[a, b]` under a delete of `count` indices at
/// `at`; `None` when it lies wholly inside the deleted span.
fn clip_interval(a: u32, b: u32, at: u32, count: u32) -> Option<(u32, u32)> {
    let end_del = at.saturating_add(count);
    if a >= at && b < end_del {
        return None;
    }
    let new_a = if a < at {
        a
    } else if a >= end_del {
        a - count
    } else {
        at
    };
    let new_b = if b < at {
        b
    } else if b >= end_del {
        b - count
    } else {
        at - 1
    };
    Some((new_a, new_b))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellProvider, Sheet};

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    /// workbook with the named sheets; caller populates cells.
    fn wb(names: &[&str]) -> Workbook {
        let mut wb = Workbook::default();
        for n in names {
            wb.sheets.push(Sheet::new(*n));
        }
        wb
    }

    fn set_formula(wb: &mut Workbook, sheet: SheetId, at: &str, f: &str) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            r(at),
            Cell {
                formula: Some(f.to_string()),
                ..Default::default()
            },
        );
    }

    fn formula(wb: &Workbook, sheet: SheetId, at: &str) -> Option<String> {
        wb.formula(sheet, r(at)).map(str::to_string)
    }

    #[test]
    fn shifts_single_cell_ref_on_insert() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "A5+1");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 3,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("A8+1"));
        assert_eq!(inv.len(), 1);
    }

    #[test]
    fn clips_range_on_delete() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "C1", "SUM(A1:A10)");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "C1").as_deref(), Some("SUM(A1:A9)"));
    }

    #[test]
    fn deleted_ref_becomes_ref_error() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "A5*2");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 4,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("#REF!*2"));
    }

    #[test]
    fn fully_deleted_range_becomes_ref_error() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "SUM(A5:A7)");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 4,
            count: 3,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("SUM(#REF!)"));
    }

    #[test]
    fn preserves_dollar_anchors() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "$A$5+1");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("$A$7+1"));
    }

    #[test]
    fn remaps_cross_sheet_ref_to_edited_sheet() {
        let mut w = wb(&["Sheet1", "Data"]);
        set_formula(&mut w, SheetId(1), "A1", "Sheet1!A5+1");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(
            formula(&w, SheetId(1), "A1").as_deref(),
            Some("Sheet1!A4+1")
        );
    }

    #[test]
    fn structural_rewrite_keeps_cell_like_sheet_name_quoted() {
        let mut workbook = wb(&["A1", "Formula"]);
        set_formula(&mut workbook, SheetId(1), "A1", "'A1'!A1");
        remap_formulas(
            &mut workbook,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            formula(&workbook, SheetId(1), "A1").as_deref(),
            Some("'A1'!A2")
        );
    }

    #[test]
    fn leaves_other_sheet_refs_untouched() {
        let mut w = wb(&["Sheet1", "Data"]);
        set_formula(&mut w, SheetId(1), "A1", "A5+1");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(1), "A1").as_deref(), Some("A5+1"));
        assert!(inv.is_empty(), "no formula changed, no inverse");
    }

    #[test]
    fn unparseable_formula_rejects_structural_rewrite() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "SUM(");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        let error = remap_formulas(&mut w, &op).unwrap_err();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("SUM("));
        assert_eq!(
            error,
            OpError::FormulaNotRewritable {
                sheet: SheetId(0),
                cell: r("B1")
            }
        );
    }

    #[test]
    fn rename_preserves_formula_source_and_unsupported_syntax() {
        let source = " SUM( Sheet1!A:A , \"Sheet1!A1\", 'SHEET1'!B2 ) ";
        assert_eq!(
            rename_formula_sheet(source, "Sheet1", "New Name"),
            " SUM( 'New Name'!A:A , \"Sheet1!A1\", 'New Name'!B2 ) "
        );
    }

    #[test]
    fn rename_escapes_quotes_in_new_sheet_name() {
        assert_eq!(
            rename_formula_sheet("Sheet1!A1", "Sheet1", "Owner's Data"),
            "'Owner''s Data'!A1"
        );
    }

    #[test]
    fn rename_handles_3d_refs_without_touching_external_or_structured_refs() {
        let source = "Sheet1:Sheet3!A1+[Book.xlsx]Sheet1!A1+Table1[Sheet1!Column]+Sheet1!A1";
        assert_eq!(
            rename_formula_sheet(source, "Sheet1", "Renamed"),
            "Renamed:Sheet3!A1+[Book.xlsx]Sheet1!A1+Table1[Sheet1!Column]+Renamed!A1"
        );
    }

    #[test]
    fn rename_rejects_formula_growth_past_the_length_cap() {
        let mut workbook = wb(&["S", "Formula"]);
        set_formula(&mut workbook, SheetId(1), "A1", "S!A1");
        let original = workbook.clone();
        let error = rename_sheet_references(&mut workbook, "S", &"x".repeat(MAX_FORMULA_BYTES))
            .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(workbook, original);
    }

    #[test]
    fn rename_quotes_cell_like_names_and_matches_unicode_sources() {
        assert_eq!(rename_formula_sheet("S!A1", "S", "A1"), "'A1'!A1");
        assert_eq!(
            rename_formula_sheet("École1!A1", "École1", "Classe"),
            "Classe!A1"
        );
        assert_eq!(
            rename_formula_sheet("Sheet😀!A1", "Sheet😀", "Renamed"),
            "Renamed!A1"
        );
        assert_eq!(rename_formula_sheet("S!A1", "S", "R4C"), "'R4C'!A1");
    }

    #[test]
    fn interval_clip_edges() {
        assert_eq!(clip_interval(0, 9, 2, 1), Some((0, 8)));
        assert_eq!(clip_interval(4, 6, 4, 3), None);
        assert_eq!(clip_interval(2, 9, 2, 4), Some((2, 5)));
        assert_eq!(clip_interval(0, 1, 5, 2), Some((0, 1)));
    }
}
