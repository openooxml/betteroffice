//! applying ops mutably to a workbook, returning the inverse for undo, plus
//! `remap_ref` — the shared address-remapping primitive.

use std::collections::BTreeMap;
use std::fmt;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{Cell, CellRange, CellRef, ColId, RowId, Sheet, SheetId, Workbook};

use crate::formatting::{mutate_number_format, patch_cell_format};
use crate::op::{CellState, Op};
use crate::remap::{
    remap_formulas, remap_hyperlink_locations, remap_hyperlink_range, rename_formula_sheet,
    rename_sheet_references,
};

/// the inverse of an applied op: a base-vocabulary op list that, replayed in
/// order, restores the prior workbook state.
#[derive(Debug, Clone, PartialEq)]
pub struct InvertedOp(pub Vec<Op>);

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OpError {
    SheetNotFound(SheetId),
    SheetIndexOutOfRange(usize),
    FormulaNotRewritable { sheet: SheetId, cell: CellRef },
    InvalidStyle(String),
}

impl fmt::Display for OpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OpError::SheetNotFound(id) => write!(f, "sheet {} not found", id.0),
            OpError::SheetIndexOutOfRange(i) => write!(f, "sheet index {i} out of range"),
            OpError::FormulaNotRewritable { sheet, cell } => write!(
                f,
                "formula at sheet {}, {} cannot be safely rewritten",
                sheet.0,
                cell.to_a1()
            ),
            OpError::InvalidStyle(message) => f.write_str(message),
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
        Op::SetFreezePane { sheet, pane } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = sheet_ref.freeze_pane;
            sheet_ref.freeze_pane = *pane;
            Ok(InvertedOp(vec![Op::SetFreezePane {
                sheet: *sheet,
                pane: old,
            }]))
        }
        Op::SetHyperlinks { sheet, hyperlinks } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut sheet_ref.hyperlinks, hyperlinks.clone());
            Ok(InvertedOp(vec![Op::SetHyperlinks {
                sheet: *sheet,
                hyperlinks: old,
            }]))
        }
        Op::MergeCells { sheet, range } => {
            let s = sheet_mut(wb, *sheet)?;
            let replaced = s
                .merges
                .iter()
                .copied()
                .filter(|merged| ranges_intersect(*merged, *range))
                .collect::<Vec<_>>();
            if replaced.as_slice() == [*range] {
                return Ok(InvertedOp(vec![]));
            }
            s.merges.retain(|merged| !ranges_intersect(*merged, *range));
            s.merges.push(*range);
            let mut inverse = Vec::with_capacity(replaced.len() + 1);
            inverse.push(Op::UnmergeCells {
                sheet: *sheet,
                range: *range,
            });
            inverse.extend(replaced.into_iter().map(|range| Op::MergeCells {
                sheet: *sheet,
                range,
            }));
            Ok(InvertedOp(inverse))
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
        Op::PatchRangeStyle {
            sheet,
            range,
            patch,
        } => apply_range_formats(wb, *sheet, *range, |format, row, col| {
            patch_cell_format(format, patch, *range, row, col)
        }),
        Op::SetRangeNumberFormat {
            sheet,
            range,
            format,
        } => apply_range_formats(wb, *sheet, *range, |cell_format, _, _| {
            mutate_number_format(cell_format, format);
            Ok(())
        }),
        Op::ApplyRangeFormat {
            sheet,
            range,
            format,
        } => {
            if format.rows == 0
                || format.columns == 0
                || format.formats.len() != (format.rows as usize) * (format.columns as usize)
            {
                return Err(OpError::InvalidStyle(
                    "captured format dimensions do not match its cells".into(),
                ));
            }
            apply_range_formats(wb, *sheet, *range, |cell_format, row, col| {
                let source_row = (row - range.start.row) % format.rows;
                let source_col = (col - range.start.col) % format.columns;
                let index = (source_row * format.columns + source_col) as usize;
                *cell_format = format.formats[index].clone();
                Ok(())
            })
        }
        Op::AddSheet { index, name } => {
            let idx = (*index).min(wb.sheets.len());
            wb.sheets.insert(idx, Sheet::new(name.clone()));
            Ok(InvertedOp(vec![Op::RemoveSheet { index: idx }]))
        }
        Op::RemoveSheet { index } => remove_sheet(wb, *index),
        Op::RenameSheet { sheet, name } => {
            let old = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let formulas = rename_sheet_references(wb, &old, name)?;
            let hyperlinks = rename_hyperlink_locations(wb, &old, name);
            sheet_mut(wb, *sheet)?.name = name.clone();
            let mut inverse = vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old,
                formulas,
            }];
            inverse.extend(hyperlinks);
            Ok(InvertedOp(inverse))
        }
        Op::RestoreSheet {
            sheet,
            name,
            formulas,
        } => {
            let old_name = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let mut old_formulas = Vec::with_capacity(formulas.len());
            for (formula_sheet, cell, _) in formulas {
                let sheet_ref = wb
                    .sheet(*formula_sheet)
                    .ok_or(OpError::SheetNotFound(*formula_sheet))?;
                let state = sheet_ref
                    .cell(*cell)
                    .map(CellState::from)
                    .unwrap_or_default();
                old_formulas.push((*formula_sheet, *cell, state));
            }
            sheet_mut(wb, *sheet)?.name = name.clone();
            for (formula_sheet, cell, state) in formulas {
                sheet_mut(wb, *formula_sheet)?.set_cell(*cell, state.clone().into());
            }
            Ok(InvertedOp(vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old_name,
                formulas: old_formulas,
            }]))
        }
        Op::InsertRows { sheet, at, count } => insert_rows(wb, *sheet, *at, *count, op),
        Op::DeleteRows { sheet, at, count } => delete_rows(wb, *sheet, *at, *count, op),
        Op::InsertCols { sheet, at, count } => insert_cols(wb, *sheet, *at, *count, op),
        Op::DeleteCols { sheet, at, count } => delete_cols(wb, *sheet, *at, *count, op),
    }
}

fn apply_range_formats(
    wb: &mut Workbook,
    sheet: SheetId,
    range: CellRange,
    mut update: impl FnMut(&mut xlsx_model::CellFormat, u32, u32) -> Result<(), OpError>,
) -> Result<InvertedOp, OpError> {
    let sheet_ref = wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let mut cells = Vec::new();
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            let at = CellRef::new(row, col);
            cells.push((at, sheet_ref.cell(at).cloned().unwrap_or_default()));
        }
    }
    let mut inverse = Vec::with_capacity(cells.len());
    for (at, old) in cells {
        let mut format = wb.styles.cell_format(old.style);
        update(&mut format, at.row, at.col)?;
        let mut next = old.clone();
        next.style = wb.styles.intern_cell_format(&format);
        if next != old {
            wb.sheet_mut(sheet)
                .ok_or(OpError::SheetNotFound(sheet))?
                .set_cell(at, next);
            inverse.push(Op::SetCell {
                sheet,
                at,
                cell: CellState::from(old),
            });
        }
    }
    Ok(InvertedOp(inverse))
}

/// apply a sequence of ops, returning the combined inverse (per-op inverses
/// concatenated in reverse order).
pub fn apply_ops(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut next = wb.clone();
    let inverse = apply_ops_in_place(&mut next, ops)?;
    *wb = next;
    Ok(inverse)
}

fn apply_ops_in_place(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
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

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && left.end.row >= right.start.row
        && left.start.col <= right.end.col
        && left.end.col >= right.start.col
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
    let restores = remap_formulas(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let dropped = shift_cells(s, op);
    shift_row_heights_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);

    let mut inv = vec![Op::DeleteRows { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let deleted = shift_cells(s, op);
    let dropped_heights = shift_row_heights_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);

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
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn insert_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let dropped = shift_cells(s, op);
    shift_col_widths_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);

    let mut inv = vec![Op::DeleteCols { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    let restores = remap_formulas(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let deleted = shift_cells(s, op);
    let dropped_widths = shift_col_widths_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);

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
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(restores);
    inv.extend(hyperlink_restores);
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

fn remap_hyperlinks(sheet: &mut Sheet, op: &Op) {
    sheet.hyperlinks = sheet
        .hyperlinks
        .drain(..)
        .filter_map(|mut hyperlink| {
            hyperlink.range = remap_hyperlink_range(hyperlink.range, op)?;
            Some(hyperlink)
        })
        .collect();
}

fn rename_hyperlink_locations(wb: &mut Workbook, old_name: &str, new_name: &str) -> Vec<Op> {
    let mut restores = Vec::new();
    for (index, sheet) in wb.sheets.iter_mut().enumerate() {
        let mut changed = false;
        let mut hyperlinks = sheet.hyperlinks.clone();
        for hyperlink in &mut hyperlinks {
            if hyperlink.external_target.is_some() {
                continue;
            }
            let Some(location) = &hyperlink.location else {
                continue;
            };
            let rewritten = rename_formula_sheet(location, old_name, new_name);
            if rewritten != *location {
                hyperlink.location = Some(rewritten);
                changed = true;
            }
        }
        if changed {
            let old = std::mem::replace(&mut sheet.hyperlinks, hyperlinks);
            restores.push(Op::SetHyperlinks {
                sheet: SheetId(index as u32),
                hyperlinks: old,
            });
        }
    }
    restores
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
    if let Some(pane) = removed.freeze_pane {
        inv.push(Op::SetFreezePane {
            sheet,
            pane: Some(pane),
        });
    }
    if !removed.hyperlinks.is_empty() {
        inv.push(Op::SetHyperlinks {
            sheet,
            hyperlinks: removed.hyperlinks,
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
    use xlsx_model::{CellProvider, CellValue, FreezePane, Hyperlink};

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
        s.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A6:B6").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
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
        assert_eq!(after.hyperlinks, orig.hyperlinks);
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
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .hyperlinks
            .push(Hyperlink {
                range: CellRange::parse_a1("A3:B3").unwrap(),
                external_target: None,
                location: Some("Sheet1!A1".into()),
                tooltip: None,
                display: None,
            });
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
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A5:B5");

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].used_range(), before);
        assert_eq!(
            wb.value(SheetId(0), r("A3")),
            CellValue::Text { value: "x".into() }
        );
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A3:B3");
    }

    #[test]
    fn structural_edits_and_renames_rewrite_hyperlink_destinations() {
        let mut wb = wb_one_sheet();
        wb.sheets[0].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1").unwrap(),
            external_target: None,
            location: Some("Target!A3".into()),
            tooltip: None,
            display: Some("Jump".into()),
        });
        wb.sheets.push(Sheet::new("Target"));
        wb.sheets[1].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1:A4").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(1),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A5")
        );
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A6");
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A3");
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A2")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(1),
                name: "New Target".into(),
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("'New Target'!A3")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);
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
        s.freeze_pane = Some(FreezePane::new(1, 1, r("C4")));
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
        assert_eq!(wb.sheets[1].freeze_pane, before.freeze_pane);
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
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1+Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inv = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
        let redo = apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb.sheets[0].name, "Sheet1");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Sheet1!A1"));
        apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
    }

    #[test]
    fn rename_undo_does_not_change_destination_only_references() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1".into()),
                ..Cell::default()
            },
        );
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A2"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        match &inverse.0[0] {
            Op::RestoreSheet { formulas, .. } => assert_eq!(formulas.len(), 1),
            other => panic!("expected restore sheet inverse, got {other:?}"),
        }
        apply_ops(&mut wb, &inverse.0).unwrap();

        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1"));
        assert_eq!(wb.formula(SheetId(1), r("A2")), Some("Sheet1!A1"));
    }

    #[test]
    fn rename_history_stays_constant_across_undo_redo() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(inverse.0.len(), 1);
        let json = serde_json::to_string(&inverse.0).unwrap();
        let decoded: Vec<Op> = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded, inverse.0);
        let redo = apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(redo.len(), 1);
        let undo = apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(undo.len(), 1);
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

    #[test]
    fn apply_ops_is_atomic_when_formula_rewriting_fails() {
        let mut wb = wb_one_sheet();
        let before = wb.clone();
        let error = apply_ops(
            &mut wb,
            &[
                Op::SetCell {
                    sheet: SheetId(0),
                    at: r("A1"),
                    cell: CellState {
                        formula: Some("SUM(".into()),
                        ..CellState::default()
                    },
                },
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
        )
        .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(wb, before);
    }
}
