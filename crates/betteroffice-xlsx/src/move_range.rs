use std::collections::BTreeMap;

use xlsx_calc::parse_formula;
use xlsx_calc::parser::Expr;
use xlsx_model::{CellRange, CellRef, SheetId, Workbook as WorkbookModel};
use xlsx_ops::{CellState, Op};

use crate::workbook::PreservedSheetState;
use crate::{Error, Result};

pub(crate) fn preserved_move_unsupported_part<'a>(
    package: &'a xlsx_parse::PreservedPackage,
    model: &WorkbookModel,
    preserved: &PreservedSheetState,
    sheet: SheetId,
    source: CellRange,
    target: CellRange,
) -> Option<&'a str> {
    let origin = preserved.origins.get(sheet.0 as usize).copied().flatten();
    let axes = preserved
        .axes
        .get(sheet.0 as usize)
        .and_then(Option::as_ref);
    let current_ref = |index, at: CellRef| {
        if Some(index) != origin {
            return None;
        }
        axes.map(|axes| {
            Some(CellRef::new(
                axes.rows.current(at.row)?,
                axes.cols.current(at.col)?,
            ))
        })
    };
    package.range_move_unsupported_part(
        |index, at| match current_ref(index, at) {
            Some(Some(at)) => source.contains(at),
            Some(None) => false,
            None => Some(index) == origin,
        },
        |index, at| match current_ref(index, at) {
            Some(Some(at)) => source.contains(at) || target.contains(at),
            Some(None) => false,
            None => Some(index) == origin,
        },
        |index, formula| {
            if !preserved.origins.contains(&Some(index)) {
                return false;
            }
            if preserved
                .axes
                .iter()
                .any(|axes| axes.as_ref().is_none_or(|axes| !axes.is_identity()))
            {
                return true;
            }
            let name = origin
                .and_then(|origin| package.source_sheet_name(origin))
                .unwrap_or(&model.sheets[sheet.0 as usize].name);
            let moved_sheet = SheetId(origin.map_or(u32::MAX, |origin| origin as u32));
            [source, target].into_iter().any(|range| {
                !matches!(
                    rewrite_formula(
                        formula,
                        SheetId(index as u32),
                        name,
                        moved_sheet,
                        range,
                        range
                    ),
                    Ok(None)
                )
            })
        },
    )
}

pub(crate) fn move_range_ops(
    model: &WorkbookModel,
    sheet: SheetId,
    source: CellRange,
    target: CellRange,
) -> Result<Vec<Op>> {
    if !model.defined_names.is_empty() {
        return Err(unsupported("defined names"));
    }
    if model.tables.iter().any(|table| {
        table.sheet == sheet && (table.range.overlaps(&source) || table.range.overlaps(&target))
    }) {
        return Err(unsupported("Excel tables"));
    }
    for current in &model.sheets {
        if !current.charts.is_empty() {
            return Err(unsupported("charts"));
        }
        if !current.hyperlinks.is_empty() {
            return Err(unsupported("hyperlinks"));
        }
    }
    let moved_sheet = model.sheet(sheet).ok_or(Error::SheetOutOfRange(sheet))?;
    if moved_sheet
        .merges
        .iter()
        .any(|range| range.overlaps(&source) || range.overlaps(&target))
    {
        return Err(unsupported("merged cells"));
    }
    if moved_sheet
        .array_formulas()
        .any(|(_, range)| range.overlaps(&source) || range.overlaps(&target))
    {
        return Err(unsupported("array formulas"));
    }

    let target_name = &moved_sheet.name;
    let mut defaults = xlsx_model::workbook::ColumnDefaults::new(moved_sheet);
    let mut desired = BTreeMap::<(u32, u32, u32), CellState>::new();
    for row in source.start.row..=source.end.row {
        for col in source.start.col..=source.end.col {
            let cleared = CellState {
                style: defaults.inherited_style(CellRef::new(row, col)).map(|_| 0),
                ..CellState::default()
            };
            desired.insert((sheet.0, row, col), cleared);
        }
    }
    for row in source.start.row..=source.end.row {
        for col in source.start.col..=source.end.col {
            let at = CellRef::new(row, col);
            let to = moved_ref(at, source, target);
            let current = moved_sheet
                .cell(at)
                .map(CellState::from)
                .unwrap_or_default();
            let mut moved = current.clone();
            moved.style = defaults
                .effective_style(at)
                .or_else(|| defaults.inherited_style(to).map(|_| 0));
            if let Some(formula) = current.formula.as_deref()
                && let Some(rewritten) =
                    rewrite_formula(formula, sheet, target_name, sheet, source, target)?
            {
                moved.formula = Some(rewritten);
            }
            desired.insert((sheet.0, to.row, to.col), moved);
        }
    }

    for (index, current_sheet) in model.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        for (at, current) in current_sheet.iter_cells() {
            if owner == sheet && (source.contains(at) || target.contains(at)) {
                continue;
            }
            let Some(formula) = current.formula.as_deref() else {
                continue;
            };
            if let Some(rewritten) =
                rewrite_formula(formula, owner, target_name, sheet, source, target)?
            {
                let mut state = CellState::from(current);
                state.formula = Some(rewritten);
                desired.insert((owner.0, at.row, at.col), state);
            }
        }
    }

    let mut ops = Vec::new();
    for ((sheet_index, row, col), state) in desired {
        let at = CellRef::new(row, col);
        let owner = SheetId(sheet_index);
        let current = model
            .sheet(owner)
            .and_then(|current| current.cell(at))
            .map(CellState::from)
            .unwrap_or_default();
        if current != state {
            ops.push(Op::SetCell {
                sheet: owner,
                at,
                cell: state,
            });
        }
    }
    Ok(ops)
}

fn rewrite_formula(
    formula: &str,
    owner: SheetId,
    target_name: &str,
    moved_sheet: SheetId,
    source: CellRange,
    target: CellRange,
) -> Result<Option<String>> {
    let expression = parse_formula(formula).map_err(|_| unsupported("unparsed formulas"))?;
    let mut changed = false;
    let rewritten = rewrite_expr(
        &expression,
        owner,
        target_name,
        moved_sheet,
        source,
        target,
        &mut changed,
    )?;
    if !changed {
        return Ok(None);
    }
    let text = rewritten.to_formula();
    parse_formula(&text).map_err(|_| unsupported("formula references"))?;
    Ok(Some(text))
}

fn rewrite_expr(
    expression: &Expr,
    owner: SheetId,
    target_name: &str,
    moved_sheet: SheetId,
    source: CellRange,
    target: CellRange,
    changed: &mut bool,
) -> Result<Expr> {
    let matches = |qualified: &Option<String>| match qualified {
        Some(name) => name.eq_ignore_ascii_case(target_name),
        None => owner == moved_sheet,
    };
    Ok(match expression {
        Expr::Ref { sheet, cell } if matches(sheet) && source.contains(*cell) => {
            *changed = true;
            Expr::Ref {
                sheet: sheet.clone(),
                cell: moved_ref(*cell, source, target),
            }
        }
        Expr::Range { sheet, range } if matches(sheet) && range.overlaps(&source) => {
            if !source.contains(range.start) || !source.contains(range.end) {
                return Err(unsupported("partially moved formula ranges"));
            }
            *changed = true;
            Expr::Range {
                sheet: sheet.clone(),
                range: CellRange {
                    start: moved_ref(range.start, source, target),
                    end: moved_ref(range.end, source, target),
                },
            }
        }
        Expr::ColumnRange { sheet, range }
            if matches(sheet) && range.start <= source.end.col && source.start.col <= range.end =>
        {
            return Err(unsupported("whole-column formula references"));
        }
        Expr::RowRange { sheet, range }
            if matches(sheet) && range.start <= source.end.row && source.start.row <= range.end =>
        {
            return Err(unsupported("whole-row formula references"));
        }
        Expr::Unary { op, expr } => Expr::Unary {
            op: *op,
            expr: Box::new(rewrite_expr(
                expr,
                owner,
                target_name,
                moved_sheet,
                source,
                target,
                changed,
            )?),
        },
        Expr::Percent(expr) => Expr::Percent(Box::new(rewrite_expr(
            expr,
            owner,
            target_name,
            moved_sheet,
            source,
            target,
            changed,
        )?)),
        Expr::Binary { op, lhs, rhs } => Expr::Binary {
            op: *op,
            lhs: Box::new(rewrite_expr(
                lhs,
                owner,
                target_name,
                moved_sheet,
                source,
                target,
                changed,
            )?),
            rhs: Box::new(rewrite_expr(
                rhs,
                owner,
                target_name,
                moved_sheet,
                source,
                target,
                changed,
            )?),
        },
        Expr::FuncCall { name, .. }
            if name.eq_ignore_ascii_case("OFFSET") || name.eq_ignore_ascii_case("INDIRECT") =>
        {
            return Err(unsupported("dynamic formula references"));
        }
        Expr::FuncCall { name, func, args } => Expr::FuncCall {
            name: name.clone(),
            func: *func,
            args: args
                .iter()
                .map(|arg| {
                    rewrite_expr(
                        arg,
                        owner,
                        target_name,
                        moved_sheet,
                        source,
                        target,
                        changed,
                    )
                })
                .collect::<Result<Vec<_>>>()?,
        },
        Expr::RangeJoin { .. } => {
            return Err(unsupported("dynamic formula ranges"));
        }
        Expr::ArrayLiteral { cols, values } => Expr::ArrayLiteral {
            cols: *cols,
            values: values
                .iter()
                .map(|value| {
                    rewrite_expr(
                        value,
                        owner,
                        target_name,
                        moved_sheet,
                        source,
                        target,
                        changed,
                    )
                })
                .collect::<Result<Vec<_>>>()?,
        },
        _ => expression.clone(),
    })
}

fn moved_ref(cell: CellRef, source: CellRange, target: CellRange) -> CellRef {
    CellRef {
        row: target.start.row + (cell.row - source.start.row),
        col: target.start.col + (cell.col - source.start.col),
        ..cell
    }
}

fn unsupported(feature: &str) -> Error {
    Error::InvalidOperation(format!("range move cannot safely preserve {feature}"))
}
