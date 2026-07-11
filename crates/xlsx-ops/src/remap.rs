//! rewriting stored formulas on row/column insert/delete: refs shift, ranges
//! clip, wholly deleted refs collapse to `#REF!`. runs before cells shift.

use std::collections::HashMap;

use xlsx_calc::parse_formula;
use xlsx_calc::parser::Expr;
use xlsx_model::{CellRange, CellRef, ErrorValue, SheetId, Workbook};

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
pub(crate) fn remap_formulas(wb: &mut Workbook, op: &Op) -> Vec<Op> {
    let Some(target) = structural_target(op) else {
        return Vec::new();
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(i, s)| (s.name.clone(), SheetId(i as u32)))
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
            let Ok(expr) = parse_formula(src) else {
                continue;
            };
            let mut changed = false;
            let new_expr = transform(&expr, op, &matches, &mut changed);
            if changed {
                restores.push(Op::SetCell {
                    sheet: owner,
                    at: cell,
                    cell: CellState::from(c),
                });
                edits.push((owner, cell, new_expr.to_formula()));
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
    restores
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
        Some(name) => names.get(name).copied(),
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
        let inv = remap_formulas(&mut w, &op);
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
        remap_formulas(&mut w, &op);
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
        remap_formulas(&mut w, &op);
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
        remap_formulas(&mut w, &op);
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
        remap_formulas(&mut w, &op);
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
        remap_formulas(&mut w, &op);
        assert_eq!(
            formula(&w, SheetId(1), "A1").as_deref(),
            Some("Sheet1!A4+1")
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
        let inv = remap_formulas(&mut w, &op);
        assert_eq!(formula(&w, SheetId(1), "A1").as_deref(), Some("A5+1"));
        assert!(inv.is_empty(), "no formula changed, no inverse");
    }

    #[test]
    fn unparseable_formula_left_verbatim() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "SUM(");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        let inv = remap_formulas(&mut w, &op);
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("SUM("));
        assert!(inv.is_empty());
    }

    #[test]
    fn interval_clip_edges() {
        assert_eq!(clip_interval(0, 9, 2, 1), Some((0, 8)));
        assert_eq!(clip_interval(4, 6, 4, 3), None);
        assert_eq!(clip_interval(2, 9, 2, 4), Some((2, 5)));
        assert_eq!(clip_interval(0, 1, 5, 2), Some((0, 1)));
    }
}
