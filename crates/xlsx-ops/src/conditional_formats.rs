use xlsx_calc::{parse_formula, parser::Expr};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellRange, CellRef, SheetId, Workbook};

use crate::remap::{formula_has_sheet_span, formula_names_sheet};
use crate::{Op, OpError};

#[derive(Clone, Copy)]
enum Edit {
    Rename {
        sheet: SheetId,
    },
    Rows {
        sheet: SheetId,
        at: u32,
    },
    Cols {
        sheet: SheetId,
        at: u32,
    },
    Move {
        sheet: SheetId,
        source: CellRange,
        target: CellRange,
    },
}

impl Edit {
    fn sheet(self) -> SheetId {
        match self {
            Self::Rename { sheet }
            | Self::Rows { sheet, .. }
            | Self::Cols { sheet, .. }
            | Self::Move { sheet, .. } => sheet,
        }
    }

    fn affects(self, range: CellRange, sqref: bool) -> bool {
        match self {
            Self::Rename { .. } => false,
            Self::Rows { at, .. } => at <= range.end.row,
            Self::Cols { at, .. } => at <= range.end.col,
            Self::Move { source, target, .. } => {
                source.overlaps(&range) || (sqref && target.overlaps(&range))
            }
        }
    }
}

/// Refuses structural ops that could affect a conditional format rule
/// authored or edited in this session. Untouched source rules keep their
/// original geometry, like data validations and auto filters.
pub fn ensure_conditional_formats_allow_edit(wb: &Workbook, op: &Op) -> Result<(), OpError> {
    let edit = match *op {
        Op::RenameSheet { sheet, ref name } => {
            let current = wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
            if current.name == *name {
                return Ok(());
            }
            Edit::Rename { sheet }
        }
        Op::InsertRows { sheet, at, count } | Op::DeleteRows { sheet, at, count } if count != 0 => {
            Edit::Rows { sheet, at }
        }
        Op::InsertCols { sheet, at, count } | Op::DeleteCols { sheet, at, count } if count != 0 => {
            Edit::Cols { sheet, at }
        }
        _ => return Ok(()),
    };
    ensure_allowed(wb, edit)
}

/// Refuses range moves that could affect a conditional format rule authored
/// or edited in this session.
pub fn ensure_conditional_formats_allow_move(
    wb: &Workbook,
    sheet: SheetId,
    source: CellRange,
    target: CellRange,
) -> Result<(), OpError> {
    if source == target {
        return Ok(());
    }
    ensure_allowed(
        wb,
        Edit::Move {
            sheet,
            source,
            target,
        },
    )
}

fn ensure_allowed(wb: &Workbook, edit: Edit) -> Result<(), OpError> {
    let target = wb
        .sheet(edit.sheet())
        .ok_or(OpError::SheetNotFound(edit.sheet()))?;
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        for group in &sheet.conditional_formats {
            let ranges = group
                .sqref
                .split_whitespace()
                .map(CellRange::parse_a1)
                .collect::<Result<Vec<_>, _>>()
                .ok()
                .filter(|ranges| !ranges.is_empty());
            for rule in group.rules.iter().filter(|rule| rule.raw_xml.is_none()) {
                let affected = if rule
                    .formulas
                    .iter()
                    .any(|formula| formula_has_sheet_span(formula))
                {
                    true
                } else if matches!(edit, Edit::Rename { .. }) {
                    rule.formulas
                        .iter()
                        .any(|formula| formula_names_sheet(formula, &target.name))
                } else if let Some(ranges) = &ranges {
                    (owner == edit.sheet() && ranges.iter().any(|range| edit.affects(*range, true)))
                        || rule.formulas.iter().any(|formula| {
                            parse_formula(formula).map_or(true, |expr| {
                                formula_affected(&expr, owner, &target.name, ranges, edit)
                            })
                        })
                } else {
                    true
                };
                if affected {
                    return Err(OpError::InvalidConditionalFormat(format!(
                        "structural edit would affect conditional format {} on sheet '{}' (sqref '{}'); reference rewriting is unsupported",
                        rule.id, sheet.name, group.sqref
                    )));
                }
            }
        }
    }
    Ok(())
}

fn formula_affected(
    expr: &Expr,
    owner: SheetId,
    target_name: &str,
    ranges: &[CellRange],
    edit: Edit,
) -> bool {
    let reference = |sheet: &Option<String>, range| {
        let matches = sheet.as_ref().map_or(owner == edit.sheet(), |name| {
            name.to_lowercase() == target_name.to_lowercase()
        });
        matches
            && ranges.iter().any(|application| {
                edit.affects(
                    relative_coverage(range, ranges[0].start, *application),
                    false,
                )
            })
    };
    let child = |expr| formula_affected(expr, owner, target_name, ranges, edit);
    match expr {
        Expr::Ref { sheet, cell } => reference(sheet, CellRange::new(*cell, *cell)),
        Expr::Range { sheet, range } => reference(sheet, *range),
        Expr::ColumnRange { sheet, range } => {
            let mut cells = range.cell_range();
            cells.start.abs_row = true;
            cells.end.abs_row = true;
            reference(sheet, cells)
        }
        Expr::RowRange { sheet, range } => {
            let mut cells = range.cell_range();
            cells.start.abs_col = true;
            cells.end.abs_col = true;
            reference(sheet, cells)
        }
        Expr::Unary { expr, .. } | Expr::Percent(expr) => child(expr),
        Expr::Binary { lhs, rhs, .. } => child(lhs) || child(rhs),
        Expr::FuncCall { name, args, .. } => {
            name.eq_ignore_ascii_case("INDIRECT")
                || name.eq_ignore_ascii_case("OFFSET")
                || args.iter().any(child)
        }
        Expr::ArrayLiteral { values, .. } => values.iter().any(child),
        Expr::RangeJoin { .. } | Expr::Name { .. } | Expr::TableRef { .. } => true,
        _ => false,
    }
}

fn relative_coverage(range: CellRange, anchor: CellRef, application: CellRange) -> CellRange {
    let shifted = |cell: CellRef, at: CellRef| {
        CellRef::new(
            if cell.abs_row {
                cell.row
            } else {
                (i64::from(cell.row) + i64::from(at.row) - i64::from(anchor.row))
                    .clamp(0, i64::from(MAX_ROWS - 1)) as u32
            },
            if cell.abs_col {
                cell.col
            } else {
                (i64::from(cell.col) + i64::from(at.col) - i64::from(anchor.col))
                    .clamp(0, i64::from(MAX_COLS - 1)) as u32
            },
        )
    };
    let first = CellRange::new(
        shifted(range.start, application.start),
        shifted(range.end, application.start),
    );
    let last = CellRange::new(
        shifted(range.start, application.end),
        shifted(range.end, application.end),
    );
    let mut coverage = CellRange::new(
        CellRef::new(
            first.start.row.min(last.start.row),
            first.start.col.min(last.start.col),
        ),
        CellRef::new(
            first.end.row.max(last.end.row),
            first.end.col.max(last.end.col),
        ),
    );
    if range.start.abs_row != range.end.abs_row {
        coverage.start.row = 0;
        coverage.end.row = MAX_ROWS - 1;
    }
    if range.start.abs_col != range.end.abs_col {
        coverage.start.col = 0;
        coverage.end.col = MAX_COLS - 1;
    }
    coverage
}
