//! applying ops mutably to a workbook, returning the inverse for undo, plus
//! `remap_ref` — the shared address-remapping primitive.

use std::collections::BTreeMap;
use std::fmt;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{Cell, CellRange, CellRef, ColId, RowId, Sheet, SheetId, Workbook};

use crate::op::{CellState, Op};
use crate::remap::remap_formulas;

/// the inverse of an applied op: a base-vocabulary op list that, replayed in
/// order, restores the prior workbook state.
#[derive(Debug, Clone, PartialEq)]
pub struct InvertedOp(pub Vec<Op>);

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OpError {
    SheetNotFound(SheetId),
    SheetIndexOutOfRange(usize),
}

impl fmt::Display for OpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OpError::SheetNotFound(id) => write!(f, "sheet {} not found", id.0),
            OpError::SheetIndexOutOfRange(i) => write!(f, "sheet index {i} out of range"),
        }
    }
}

impl std::error::Error for OpError {}

/// apply one op, mutating `wb` and returning its inverse.
pub fn apply(wb: &mut Workbook, op: &Op) -> Result<InvertedOp, OpError> {
    match op {
        Op::SetCell { sheet, at, cell } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.cell(*at).map(CellState::from).unwrap_or_default();
            s.set_cell(*at, cell.clone().into());
            Ok(InvertedOp(vec![Op::SetCell {
                sheet: *sheet,
                at: *at,
                cell: old,
            }]))
        }
        Op::SetColWidth { sheet, col, width } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.col_widths.get(col).copied();
            match width {
                Some(w) => {
                    s.col_widths.insert(*col, *w);
                }
                None => {
                    s.col_widths.remove(col);
                }
            }
            Ok(InvertedOp(vec![Op::SetColWidth {
                sheet: *sheet,
                col: *col,
                width: old,
            }]))
        }
        Op::SetRowHeight { sheet, row, height } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.row_heights.get(row).copied();
            match height {
                Some(h) => {
                    s.row_heights.insert(*row, *h);
                }
                None => {
                    s.row_heights.remove(row);
                }
            }
            Ok(InvertedOp(vec![Op::SetRowHeight {
                sheet: *sheet,
                row: *row,
                height: old,
            }]))
        }
        Op::MergeCells { sheet, range } => {
            let s = sheet_mut(wb, *sheet)?;
            if s.merges.contains(range) {
                return Ok(InvertedOp(vec![]));
            }
            s.merges.push(*range);
            Ok(InvertedOp(vec![Op::UnmergeCells {
                sheet: *sheet,
                range: *range,
            }]))
        }
        Op::UnmergeCells { sheet, range } => {
            let s = sheet_mut(wb, *sheet)?;
            match s.merges.iter().position(|m| m == range) {
                Some(pos) => {
                    s.merges.remove(pos);
                    Ok(InvertedOp(vec![Op::MergeCells {
                        sheet: *sheet,
                        range: *range,
                    }]))
                }
                None => Ok(InvertedOp(vec![])),
            }
        }
        Op::AddSheet { index, name } => {
            let idx = (*index).min(wb.sheets.len());
            wb.sheets.insert(idx, Sheet::new(name.clone()));
            Ok(InvertedOp(vec![Op::RemoveSheet { index: idx }]))
        }
        Op::RemoveSheet { index } => remove_sheet(wb, *index),
        Op::RenameSheet { sheet, name } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut s.name, name.clone());
            Ok(InvertedOp(vec![Op::RenameSheet {
                sheet: *sheet,
                name: old,
            }]))
        }
        Op::InsertRows { sheet, at, count } => insert_rows(wb, *sheet, *at, *count, op),
        Op::DeleteRows { sheet, at, count } => delete_rows(wb, *sheet, *at, *count, op),
        Op::InsertCols { sheet, at, count } => insert_cols(wb, *sheet, *at, *count, op),
        Op::DeleteCols { sheet, at, count } => delete_cols(wb, *sheet, *at, *count, op),
    }
}

/// apply a sequence of ops, returning the combined inverse (per-op inverses
/// concatenated in reverse order).
pub fn apply_ops(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut per_op: Vec<Vec<Op>> = Vec::with_capacity(ops.len());
    for op in ops {
        per_op.push(apply(wb, op)?.0);
    }
    let mut inverse = Vec::new();
    for chunk in per_op.into_iter().rev() {
        inverse.extend(chunk);
    }
    Ok(inverse)
}

/// remap an address through a structural op; `None` when it falls inside a
/// deleted span. non-structural ops leave the address unchanged.
pub fn remap_ref(at: CellRef, structural_op: &Op) -> Option<CellRef> {
    match *structural_op {
        Op::InsertRows { at: p, count, .. } => {
            shift_row(at, |r| insert_axis(r, p, count, MAX_ROWS))
        }
        Op::DeleteRows { at: p, count, .. } => shift_row(at, |r| delete_axis(r, p, count)),
        Op::InsertCols { at: p, count, .. } => {
            shift_col(at, |c| insert_axis(c, p, count, MAX_COLS))
        }
        Op::DeleteCols { at: p, count, .. } => shift_col(at, |c| delete_axis(c, p, count)),
        _ => Some(at),
    }
}

/// remap both corners of a range; `None` if either corner is deleted.
fn remap_range(range: CellRange, op: &Op) -> Option<CellRange> {
    let start = remap_ref(range.start, op)?;
    let end = remap_ref(range.end, op)?;
    Some(CellRange::new(start, end))
}

/// index shift under an insert at `p` of `count`. indices at or after `p` move
/// up by `count`; anything pushed past `limit` is dropped (`None`).
fn insert_axis(idx: u32, p: u32, count: u32, limit: u32) -> Option<u32> {
    if idx < p {
        return Some(idx);
    }
    let shifted = u64::from(idx) + u64::from(count);
    if shifted > u64::from(limit - 1) {
        None
    } else {
        Some(shifted as u32)
    }
}

/// index shift under a delete at `p` of `count`. indices inside `[p, p+count)`
/// are deleted (`None`); later indices move down by `count`.
fn delete_axis(idx: u32, p: u32, count: u32) -> Option<u32> {
    if idx < p {
        Some(idx)
    } else if idx < p.saturating_add(count) {
        None
    } else {
        Some(idx - count)
    }
}

fn shift_row(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.row = f(at.row)?;
    Some(at)
}

fn shift_col(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.col = f(at.col)?;
    Some(at)
}

fn insert_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let dropped = shift_cells(s, op);
    shift_row_heights_up(s, at, count);
    remap_merges_keep(s, op);

    let mut inv = vec![Op::DeleteRows { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.extend(restores);
    Ok(InvertedOp(inv))
}

fn delete_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let deleted = shift_cells(s, op);
    let dropped_heights = shift_row_heights_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);

    let mut inv = vec![Op::InsertRows { sheet, at, count }];
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (row, h) in dropped_heights {
        inv.push(Op::SetRowHeight {
            sheet,
            row,
            height: Some(h),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.extend(restores);
    Ok(InvertedOp(inv))
}

fn insert_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let dropped = shift_cells(s, op);
    shift_col_widths_up(s, at, count);
    remap_merges_keep(s, op);

    let mut inv = vec![Op::DeleteCols { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.extend(restores);
    Ok(InvertedOp(inv))
}

fn delete_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let deleted = shift_cells(s, op);
    let dropped_widths = shift_col_widths_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);

    let mut inv = vec![Op::InsertCols { sheet, at, count }];
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (col, w) in dropped_widths {
        inv.push(Op::SetColWidth {
            sheet,
            col,
            width: Some(w),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.extend(restores);
    Ok(InvertedOp(inv))
}

/// remap every occupied cell through `op`, rebuilding storage. returns the
/// cells whose address was dropped, for the inverse.
fn shift_cells(s: &mut Sheet, op: &Op) -> Vec<(CellRef, Cell)> {
    let old: Vec<(CellRef, Cell)> = s.iter_cells().map(|(r, c)| (r, c.clone())).collect();
    let mut moved = Vec::new();
    let mut dropped = Vec::new();
    for (r, c) in &old {
        match remap_ref(*r, op) {
            Some(nr) => moved.push((nr, c.clone())),
            None => dropped.push((*r, c.clone())),
        }
    }
    for (r, _) in &old {
        s.set_cell(*r, Cell::default());
    }
    for (nr, c) in moved {
        s.set_cell(nr, c);
    }
    dropped
}

fn shift_row_heights_up(s: &mut Sheet, at: RowId, count: u32) {
    let shifted: BTreeMap<RowId, f64> = s
        .row_heights
        .iter()
        .map(|(&row, &h)| {
            (
                if row >= at {
                    row.saturating_add(count)
                } else {
                    row
                },
                h,
            )
        })
        .collect();
    s.row_heights = shifted;
}

fn shift_row_heights_down(s: &mut Sheet, at: RowId, count: u32) -> Vec<(RowId, f64)> {
    let mut kept = BTreeMap::new();
    let mut dropped = Vec::new();
    for (&row, &h) in &s.row_heights {
        if row < at {
            kept.insert(row, h);
        } else if row >= at.saturating_add(count) {
            kept.insert(row - count, h);
        } else {
            dropped.push((row, h));
        }
    }
    s.row_heights = kept;
    dropped
}

fn shift_col_widths_up(s: &mut Sheet, at: ColId, count: u32) {
    let shifted: BTreeMap<ColId, f64> = s
        .col_widths
        .iter()
        .map(|(&col, &w)| {
            (
                if col >= at {
                    col.saturating_add(count)
                } else {
                    col
                },
                w,
            )
        })
        .collect();
    s.col_widths = shifted;
}

fn shift_col_widths_down(s: &mut Sheet, at: ColId, count: u32) -> Vec<(ColId, f64)> {
    let mut kept = BTreeMap::new();
    let mut dropped = Vec::new();
    for (&col, &w) in &s.col_widths {
        if col < at {
            kept.insert(col, w);
        } else if col >= at.saturating_add(count) {
            kept.insert(col - count, w);
        } else {
            dropped.push((col, w));
        }
    }
    s.col_widths = kept;
    dropped
}

/// remap merges under an insert (no corner is ever deleted).
fn remap_merges_keep(s: &mut Sheet, op: &Op) {
    let remapped: Vec<CellRange> = s
        .merges
        .iter()
        .filter_map(|m| remap_range(*m, op))
        .collect();
    s.merges = remapped;
}

/// remap merges under a delete; any merge with a deleted corner is dropped and
/// returned for the inverse to restore.
fn remap_merges_drop(s: &mut Sheet, op: &Op) -> Vec<CellRange> {
    let mut kept = Vec::new();
    let mut dropped = Vec::new();
    for m in &s.merges {
        match remap_range(*m, op) {
            Some(nm) => kept.push(nm),
            None => dropped.push(*m),
        }
    }
    s.merges = kept;
    dropped
}

fn remove_sheet(wb: &mut Workbook, index: usize) -> Result<InvertedOp, OpError> {
    if index >= wb.sheets.len() {
        return Err(OpError::SheetIndexOutOfRange(index));
    }
    let removed = wb.sheets.remove(index);
    let sheet = SheetId(index as u32);
    let mut inv = vec![Op::AddSheet {
        index,
        name: removed.name.clone(),
    }];
    for (at, cell) in removed.iter_cells() {
        inv.push(Op::SetCell {
            sheet,
            at,
            cell: CellState::from(cell),
        });
    }
    for range in &removed.merges {
        inv.push(Op::MergeCells {
            sheet,
            range: *range,
        });
    }
    for (&col, &w) in &removed.col_widths {
        inv.push(Op::SetColWidth {
            sheet,
            col,
            width: Some(w),
        });
    }
    for (&row, &h) in &removed.row_heights {
        inv.push(Op::SetRowHeight {
            sheet,
            row,
            height: Some(h),
        });
    }
    Ok(InvertedOp(inv))
}

fn sheet_mut(wb: &mut Workbook, sheet: SheetId) -> Result<&mut Sheet, OpError> {
    wb.sheet_mut(sheet).ok_or(OpError::SheetNotFound(sheet))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{CellProvider, CellValue};

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    fn num(v: f64) -> CellState {
        CellState {
            value: CellValue::Number { value: v },
            ..Default::default()
        }
    }

    fn wb_one_sheet() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb
    }

    #[test]
    fn remap_ref_insert_rows_edges() {
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        // a1 row N is 0-based index N-1, so A5 sits before the insert at 5
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(r("A6"), &op).unwrap().row, 5 + 3);
        assert_eq!(remap_ref(r("A100"), &op).unwrap().row, 99 + 3);
    }

    #[test]
    fn remap_ref_delete_rows_edges() {
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(CellRef::new(5, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(7, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(8, 0), &op), Some(CellRef::new(5, 0)));
    }

    #[test]
    fn remap_ref_cols_and_anchors() {
        let ins = Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        };
        let mut anchored = r("C1");
        anchored.abs_col = true;
        let out = remap_ref(anchored, &ins).unwrap();
        assert_eq!(out.col, 3);
        assert!(out.abs_col, "anchor preserved through remap");

        let del = Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        assert_eq!(remap_ref(CellRef::new(0, 2), &del), None);
        assert_eq!(
            remap_ref(CellRef::new(0, 4), &del),
            Some(CellRef::new(0, 2))
        );
    }

    #[test]
    fn remap_ref_ignores_non_structural() {
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("A1"),
            cell: num(1.0),
        };
        assert_eq!(remap_ref(r("Z9"), &op), Some(r("Z9")));
    }

    #[test]
    fn set_cell_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("B2"),
            cell: num(42.0),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("B2")),
            CellValue::Number { value: 42.0 }
        );
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.value(SheetId(0), r("B2")), CellValue::Empty);
        assert!(wb.sheet(SheetId(0)).unwrap().used_range().is_none());
    }

    #[test]
    fn delete_rows_restores_contents_and_merges() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        s.set_cell(
            r("A6"),
            Cell {
                value: CellValue::Number { value: 6.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A6:B6").unwrap());
        s.row_heights.insert(5, 30.0);
        let before = wb.clone();

        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A4")),
            CellValue::Number { value: 6.0 }
        );

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        let after = &wb.sheets[0];
        let orig = &before.sheets[0];
        assert_eq!(after.used_range(), orig.used_range());
        assert_eq!(
            wb.value(SheetId(0), r("A6")),
            CellValue::Number { value: 6.0 }
        );
        assert_eq!(after.merges, orig.merges);
        assert_eq!(after.row_heights, orig.row_heights);
    }

    #[test]
    fn insert_rows_round_trip() {
        let mut wb = wb_one_sheet();
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            r("A3"),
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Default::default()
            },
        );
        let before = wb.sheets[0].used_range();

        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A5")),
            CellValue::Text { value: "x".into() }
        );

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].used_range(), before);
        assert_eq!(
            wb.value(SheetId(0), r("A3")),
            CellValue::Text { value: "x".into() }
        );
    }

    #[test]
    fn add_and_remove_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::AddSheet {
            index: 1,
            name: "New".into(),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[1].name, "New");
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 1);
    }

    #[test]
    fn remove_populated_sheet_round_trips_contents() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Two"));
        let s = wb.sheet_mut(SheetId(1)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 9.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A1:B1").unwrap());
        s.col_widths.insert(0, 12.0);
        let before = wb.sheets[1].clone();

        let inv = apply(&mut wb, &Op::RemoveSheet { index: 1 }).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(
            wb.value(SheetId(1), r("A1")),
            CellValue::Number { value: 9.0 }
        );
        assert_eq!(wb.sheets[1].merges, before.merges);
        assert_eq!(wb.sheets[1].col_widths, before.col_widths);
    }

    #[test]
    fn merge_and_col_width_round_trip() {
        let mut wb = wb_one_sheet();
        let range = CellRange::parse_a1("A1:B2").unwrap();
        let inv = apply(
            &mut wb,
            &Op::MergeCells {
                sheet: SheetId(0),
                range,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].merges, vec![range]);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert!(wb.sheets[0].merges.is_empty());

        let inv = apply(
            &mut wb,
            &Op::SetColWidth {
                sheet: SheetId(0),
                col: 3,
                width: Some(20.0),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].col_widths.get(&3), Some(&20.0));
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].col_widths.get(&3), None);
    }

    #[test]
    fn rename_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        let inv = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].name, "Renamed");
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].name, "Sheet1");
    }

    #[test]
    fn missing_sheet_errors() {
        let mut wb = Workbook::default();
        let err = apply(
            &mut wb,
            &Op::SetCell {
                sheet: SheetId(3),
                at: r("A1"),
                cell: num(1.0),
            },
        );
        assert_eq!(err.unwrap_err(), OpError::SheetNotFound(SheetId(3)));
    }
}
