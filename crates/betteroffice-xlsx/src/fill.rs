//! Copy-fill for the grid fill handle: tile the source block over the
//! destination and edit the copies as one undo step.
//!
//! This is COPY tiling only. Values (including a formula's cached result) and
//! style indices are copied; formulas are translated from the cell in the tile
//! they came from to the cell they land on, so relative references shift and
//! `$` references stay put. There is no series/trend detection and no clearing
//! gesture: a pointer above or left of the source simply extends nothing.
//!
//! Everything a tiled copy cannot faithfully preserve is refused before any
//! op is built, so a refusal leaves the workbook untouched.

use xlsx_model::workbook::ColumnDefaults;
use xlsx_model::{CellFormat, CellRange, CellRef, SheetId, Workbook as WorkbookModel};
use xlsx_ops::{CellState, Op};

use crate::workbook::PreservedSheetState;
use crate::{Error, Result};

fn refuses(feature: &str) -> Error {
    Error::InvalidOperation(format!("fill cannot safely preserve {feature}"))
}

fn covers_partially(ranges: &[CellRange], area: CellRange) -> bool {
    let mut covered = false;
    let mut uncovered = false;
    for row in area.start.row..=area.end.row {
        for col in area.start.col..=area.end.col {
            if ranges.iter().any(|range| range.contains(CellRef::new(row, col))) {
                covered = true;
            } else {
                uncovered = true;
            }
            if covered && uncovered {
                return true;
            }
        }
    }
    false
}

fn rule_ranges(sqref: &str) -> Option<Vec<CellRange>> {
    let ranges = sqref
        .split_whitespace()
        .map(CellRange::parse_a1)
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    (!ranges.is_empty()).then_some(ranges)
}

/// The cells `destination` fills, each as the source cell whose tile it repeats.
fn tiled_source(source: CellRange, row: u32, col: u32) -> CellRef {
    let rows = source.end.row - source.start.row + 1;
    let cols = source.end.col - source.start.col + 1;
    CellRef::new(
        source.start.row + (row - source.start.row) % rows,
        source.start.col + (col - source.start.col) % cols,
    )
}

/// Refuses a fill over the data validations and comments a source sheet part
/// keeps. Once rows or columns moved since opening, any validation refuses.
pub(crate) fn preserved_fill_refusal(
    package: &xlsx_parse::PreservedPackage,
    preserved: &PreservedSheetState,
    sheet: SheetId,
    destination: CellRange,
) -> Option<Error> {
    let origin = preserved.origins.get(sheet.0 as usize).copied().flatten()?;
    let axes = preserved
        .axes
        .get(sheet.0 as usize)
        .and_then(Option::as_ref);
    let exact = axes.is_some_and(xlsx_parse::SheetAxes::is_identity);
    let feature = package.fill_unsupported_feature(
        origin,
        |ranges| !exact || covers_partially(ranges, destination),
        |at| {
            axes.and_then(|axes| {
                Some(CellRef::new(
                    axes.rows.current(at.row)?,
                    axes.cols.current(at.col)?,
                ))
            })
            .is_none_or(|at| destination.contains(at))
        },
    )?;
    Some(refuses(feature))
}

/// Builds the `SetCell` ops that tile `source` over `destination`.
///
/// `destination` is the complete filled rectangle and starts at the source's
/// top-left corner, so it contains the source; cells inside the source are left
/// alone, and every other cell in the rectangle is copied from its tile.
pub(crate) fn fill_range_ops(
    model: &WorkbookModel,
    sheet: SheetId,
    source: CellRange,
    destination: CellRange,
) -> Result<Vec<Op>> {
    if source.start != destination.start
        || destination.end.row < source.end.row
        || destination.end.col < source.end.col
    {
        return Err(Error::InvalidOperation(
            "fill destination must start at the source's top-left cell and contain it".to_string(),
        ));
    }
    let target = model.sheet(sheet).ok_or(Error::SheetOutOfRange(sheet))?;
    if target
        .array_formulas()
        .any(|(_, spill)| spill.overlaps(&source) || spill.overlaps(&destination))
    {
        return Err(refuses("array formulas"));
    }
    if target
        .merges
        .iter()
        .any(|merged| merged.overlaps(&source) || merged.overlaps(&destination))
    {
        return Err(refuses("merged cells"));
    }
    if model.tables.iter().any(|table| {
        table.sheet == sheet
            && (table.range.overlaps(&source) || table.range.overlaps(&destination))
    }) {
        return Err(refuses("Excel tables"));
    }
    if target.charts.iter().any(|chart| {
        chart.anchor.from_cell().is_some_and(|anchor| {
            let at = CellRef::new(anchor.row, anchor.col);
            source.contains(at) || destination.contains(at)
        })
    }) {
        return Err(refuses("charts anchored in the filled range"));
    }
    if target
        .hyperlinks
        .iter()
        .any(|link| link.range.overlaps(&source) || link.range.overlaps(&destination))
    {
        return Err(refuses("hyperlinks"));
    }
    for group in &target.conditional_formats {
        match rule_ranges(&group.sqref) {
            Some(ranges) if covers_partially(&ranges, destination) => {
                return Err(refuses("conditional formats"));
            }
            None if !group.sqref.trim().is_empty() => {
                return Err(refuses("conditional formats"));
            }
            _ => {}
        }
    }

    let rows = (source.end.row - source.start.row + 1) as usize;
    let cols = (source.end.col - source.start.col + 1) as usize;
    let mut defaults = ColumnDefaults::new(target);
    let mut tile = Vec::with_capacity(rows * cols);
    for row in source.start.row..=source.end.row {
        for col in source.start.col..=source.end.col {
            let at = CellRef::new(row, col);
            let state = match target.cell(at) {
                Some(cell) => CellState::from(cell),
                None => CellState {
                    // an absent source cell displays its row/column default,
                    // which the fill carries to the cells it writes.
                    style: defaults.effective_style(at),
                    ..CellState::default()
                },
            };
            tile.push(state);
        }
    }

    let mut ops = Vec::new();
    for row in destination.start.row..=destination.end.row {
        for col in destination.start.col..=destination.end.col {
            let at = CellRef::new(row, col);
            if source.contains(at) {
                continue;
            }
            let origin = tiled_source(source, row, col);
            let index = (origin.row - source.start.row) as usize * cols
                + (origin.col - source.start.col) as usize;
            let mut state = tile[index].clone();
            if let Some(formula) = state.formula.take() {
                let translated =
                    xlsx_parse::translate_formula(&formula, origin, at).map_err(|_| {
                        Error::InvalidOperation(
                            "fill cannot copy a formula above the length limit".to_string(),
                        )
                    })?;
                state.formula = Some(translated);
            }
            let current = target.cell(at).map(CellState::from).unwrap_or_default();
            if current != state {
                ops.push(Op::SetCell {
                    sheet,
                    at,
                    cell: state.clone(),
                });
            }
            // A copied blank appearance must override a formatted row or
            // column. Applying the plain format through an op materializes a
            // plain xf when the workbook has none and records it in the edit,
            // so the pin has no effect on refusals and saves/undoes with the
            // rest of the fill.
            let target_style = if state == CellState::default() {
                defaults.inherited_style(at)
            } else {
                state.style
            };
            let plain = model.styles.resolved_format(None);
            if state.style.is_none()
                && target_style
                    .is_some_and(|style| model.styles.resolved_format(Some(style)) != plain)
            {
                ops.push(Op::ApplyRangeFormat {
                    sheet,
                    range: CellRange::new(at, at),
                    format: crate::CapturedFormat {
                        rows: 1,
                        columns: 1,
                        formats: vec![CellFormat::default()],
                        kept: Vec::new(),
                        source_styles: Vec::new(),
                        source_formats: Vec::new(),
                    },
                });
            }
        }
    }
    Ok(ops)
}
