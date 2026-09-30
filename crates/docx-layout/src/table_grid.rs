//! DOCX table-grid geometry and column-width resolution.
//!
//! [`resolve_cell_grid`] is the single source of truth for which grid column a
//! cell occupies: measurement, painting and the row-break paginator all read it,
//! so they cannot disagree. It is deliberately width-free — callers multiply a
//! column index by their own (possibly scaled) widths.
//!
//! Column widths come from one of three algorithms, chosen by
//! `w:tblLayout`/the block's width algorithm:
//!
//! - **fixed** — normalize the declared grid, then apply the first row's
//!   preferred cell widths before fitting any explicit table width.
//! - **autofit** — accumulate per-column minimum and maximum content widths
//!   from every cell, bounded by preferred cell widths and the content budget.
//!   Distribute room above minimums in proportion to column flex (`max - min`),
//!   or shrink minimums proportionally when they exceed the budget.
//! - otherwise — normalize the declared grid and uniformly scale it to an
//!   explicit table width when the two differ by more than a pixel. A table
//!   that declares no width of its own then lets [`content_sized_columns`] and
//!   [`grow_content_sized_columns`] widen the columns no cell prices, which is
//!   where Word re-measures and the stored `w:gridCol` goes stale.
//!
//! A width pair resolves through the preferred-width element, then the flat
//! value/type pair, then a raw pixel width. `pct` units are 50ths of a percent
//! (ECMA-376 §17.18.111, so 5000 means 100%); `dxa` and an absent type are twips
//! at 96 DPI. Automatic, zero, negative and NaN widths never resolve.

use std::collections::{HashMap, HashSet};

pub use ooxml_drawingml::normalize_table_column_widths;

use serde::Serialize;

use crate::measure_blocks::DEFAULT_CELL_PADDING_X;
use crate::types::TableBlock;

/// Twips per inch.
const TWIPS_PER_INCH: f64 = 1440.0;
/// Pixels per inch at the standard 96 DPI assumption.
const PIXELS_PER_INCH: f64 = 96.0;

/// Converts twips to pixels at 96 DPI without reassociating arithmetic.
fn twips_to_pixels(twips: f64) -> f64 {
    (twips / TWIPS_PER_INCH) * PIXELS_PER_INCH
}

/// JS truthiness for a number: `0`, `-0` and `NaN` are falsy.
fn js_truthy(v: f64) -> bool {
    v != 0.0 && !v.is_nan()
}

/// Resolves twips or a percentage in 50ths of a percent to pixels.
pub fn resolve_table_width_px(
    value: Option<f64>,
    width_type: Option<&str>,
    parent_width: f64,
) -> Option<f64> {
    let value = value?;
    if !(value > 0.0) {
        return None;
    }
    if width_type == Some("pct") {
        return Some((parent_width * value) / 5000.0);
    }
    if width_type.is_none() || width_type == Some("dxa") {
        return Some(twips_to_pixels(value));
    }
    None
}

/// A cell with its resolved grid position (column index honoring spans).
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedGridCell {
    pub row_index: usize,
    pub cell_index: usize,
    pub column_index: usize,
    pub col_span: usize,
    pub row_span: usize,
}

/// Resolve every cell's grid column index.
///
/// A row starts at its `w:gridBefore` offset and skips any column a
/// vertically-merged cell from an earlier row still occupies; each cell then
/// claims `w:gridSpan` columns, or starts at its explicit grid position when
/// one is given. Spans are truncated and clamped (columns to 16384 rows to
/// 32768), so malformed input cannot blow up the grid.
pub fn resolve_cell_grid(table_block: &TableBlock) -> Vec<ResolvedGridCell> {
    let mut occupied: HashMap<usize, HashSet<usize>> = HashMap::new();
    let mut out: Vec<ResolvedGridCell> = Vec::new();
    for row_index in 0..table_block.rows.len() {
        let cells = &table_block.rows[row_index].cells;
        // Rows only ever seed sets for LATER rows, so taking ownership of this
        // row's set is safe.
        let occ = occupied.remove(&row_index).unwrap_or_default();
        let mut column_index = table_block.rows[row_index]
            .grid_before
            .unwrap_or(0)
            .min(16_384) as usize;
        while occ.contains(&column_index) {
            column_index += 1;
        }
        for (cell_index, cell) in cells.iter().enumerate() {
            if let Some(grid_start) = cell.grid_start {
                column_index = column_index.max(grid_start.min(16_384) as usize);
            }
            let col_span = (cell.col_span.unwrap_or(1.0).trunc() as usize).clamp(1, 16_384);
            let row_span = (cell.row_span.unwrap_or(1.0).trunc() as usize).clamp(1, 32_768);
            out.push(ResolvedGridCell {
                row_index,
                cell_index,
                column_index,
                col_span,
                row_span,
            });
            if row_span > 1 {
                for r in row_index + 1..row_index + row_span {
                    let s = occupied.entry(r).or_default();
                    for c in 0..col_span {
                        s.insert(column_index + c);
                    }
                }
            }
            column_index += col_span;
            while occ.contains(&column_index) {
                column_index += 1;
            }
        }
    }
    out
}

/// Total grid columns: the furthest column any cell reaches, and every row's
/// own end plus its `w:gridAfter` skip.
pub fn count_table_columns(table_block: &TableBlock) -> usize {
    let resolved = resolve_cell_grid(table_block);
    let mut count = 1usize;
    for cell in &resolved {
        count = count.max(cell.column_index + cell.col_span);
    }
    for (row_index, row) in table_block.rows.iter().enumerate() {
        let row_end = resolved
            .iter()
            .filter(|cell| cell.row_index == row_index)
            .map(|cell| cell.column_index + cell.col_span)
            .fold(row.grid_before.unwrap_or(0) as usize, usize::max);
        count = count.max(row_end + row.grid_after.unwrap_or(0) as usize);
    }
    count.min(16_384)
}

/// Resolves a width through the preferred element, then the flat value/type
/// pair, then a raw pixel width.
pub(crate) fn preferred_width_px(
    preferred: Option<&crate::types::PreferredWidth>,
    legacy_value: Option<f64>,
    legacy_type: Option<&str>,
    parent_width: f64,
    legacy_px: Option<f64>,
) -> Option<f64> {
    if preferred.is_some_and(|width| width.r#type.as_deref() == Some("auto")) {
        return None;
    }
    preferred
        .and_then(|width| {
            resolve_table_width_px(width.value, width.r#type.as_deref(), parent_width)
        })
        .or_else(|| {
            if legacy_type == Some("auto") {
                return None;
            }
            resolve_table_width_px(legacy_value, legacy_type, parent_width)
                .or_else(|| legacy_px.filter(|value| *value > 0.0))
        })
}

/// Raises a span's columns until they total `required`, sharing the shortfall
/// evenly. Columns already wide enough are left alone.
fn add_span_constraint(widths: &mut [f64], start: usize, span: usize, required: f64) {
    if !(required > 0.0) || start >= widths.len() {
        return;
    }
    let end = widths.len().min(start + span.max(1));
    let current: f64 = widths[start..end].iter().sum();
    let deficit = required - current;
    if deficit <= 0.0 {
        return;
    }
    let share = deficit / (end - start).max(1) as f64;
    for width in &mut widths[start..end] {
        *width += share;
    }
}

/// Grows every column equally up to `target`. Never shrinks.
fn distribute_to_target(mut widths: Vec<f64>, target: f64) -> Vec<f64> {
    let current: f64 = widths.iter().sum();
    if target > current && !widths.is_empty() {
        let share = (target - current) / widths.len() as f64;
        for width in &mut widths {
            *width += share;
        }
    }
    widths
}

/// Fixed layout: only the first row's cells constrain the grid.
fn resolve_fixed_column_widths(
    table_block: &TableBlock,
    content_width: f64,
    col_count: usize,
    explicit_width_px: Option<f64>,
) -> Vec<f64> {
    let source = table_block
        .grid_widths
        .as_deref()
        .or(table_block.column_widths.as_deref())
        .unwrap_or(&[]);
    let mut widths = normalize_table_column_widths(
        source,
        col_count,
        explicit_width_px.unwrap_or(content_width),
    );
    let table_width = explicit_width_px.unwrap_or_else(|| widths.iter().sum());
    let mut has_cell_preferences = false;
    for grid_cell in resolve_cell_grid(table_block)
        .into_iter()
        .filter(|cell| cell.row_index == 0)
    {
        let Some(cell) = table_block.rows[0].cells.get(grid_cell.cell_index) else {
            continue;
        };
        if let Some(preferred) = preferred_width_px(
            cell.preferred_width.as_ref(),
            cell.width_value,
            cell.width_type.as_deref(),
            table_width,
            cell.width,
        ) {
            let start = grid_cell.column_index;
            let end = widths.len().min(start + grid_cell.col_span);
            if start >= end {
                continue;
            }
            has_cell_preferences = true;
            let total: f64 = widths[start..end].iter().sum();
            if total > preferred {
                let scale = preferred / total;
                for width in &mut widths[start..end] {
                    *width *= scale;
                }
            } else {
                add_span_constraint(&mut widths, start, grid_cell.col_span, preferred);
            }
        }
    }
    if let Some(target) = explicit_width_px {
        let total: f64 = widths.iter().sum();
        if total > target || (!has_cell_preferences && total > 0.0) {
            let scale = target / total;
            return widths.into_iter().map(|width| width * scale).collect();
        }
        return distribute_to_target(widths, target);
    }
    widths
}

fn autofit_content_widths(
    table_block: &TableBlock,
    content_width: f64,
    col_count: usize,
    explicit_width_px: Option<f64>,
    content_widths: Option<&[Vec<Option<(f64, f64)>>]>,
) -> (Vec<f64>, Vec<f64>) {
    let source = table_block
        .grid_widths
        .as_deref()
        .or(table_block.column_widths.as_deref())
        .unwrap_or(&[]);
    let base = normalize_table_column_widths(
        source,
        col_count,
        explicit_width_px.unwrap_or(content_width),
    );
    let mut minimums = vec![0.0; col_count];
    let mut constraints = Vec::new();
    for grid_cell in resolve_cell_grid(table_block) {
        let Some(cell) = table_block
            .rows
            .get(grid_cell.row_index)
            .and_then(|row| row.cells.get(grid_cell.cell_index))
        else {
            continue;
        };
        let preferred = preferred_width_px(
            cell.preferred_width.as_ref(),
            cell.width_value,
            cell.width_type.as_deref(),
            explicit_width_px.unwrap_or(0.0),
            cell.width,
        )
        .filter(|width| *width > 0.0);
        let width_type = cell
            .preferred_width
            .as_ref()
            .filter(|width| {
                width.r#type.as_deref() == Some("auto")
                    || resolve_table_width_px(width.value, width.r#type.as_deref(), 1.0).is_some()
            })
            .map(|width| width.r#type.as_deref())
            .unwrap_or(cell.width_type.as_deref());
        let percentage = if width_type == Some("pct") {
            preferred_width_px(
                cell.preferred_width.as_ref(),
                cell.width_value,
                cell.width_type.as_deref(),
                1.0,
                cell.width,
            )
        } else {
            None
        };
        let measured = content_widths
            .and_then(|rows| rows.get(grid_cell.row_index))
            .and_then(|cells| cells.get(grid_cell.cell_index))
            .copied()
            .flatten();
        let max_content_width = cell.max_content_width.or(measured.map(|widths| widths.1));
        let mut minimum = cell
            .min_content_width
            .or(measured.map(|widths| widths.0))
            .unwrap_or(0.0)
            .max(0.0);
        if cell.no_wrap.unwrap_or(false) {
            let absolute =
                width_type == Some("dxa") || (width_type.is_none() && preferred.is_some());
            minimum = minimum.max(if absolute {
                preferred.unwrap_or(0.0)
            } else {
                max_content_width.unwrap_or(0.0)
            });
        }
        let maximum = minimum.max(preferred.or(max_content_width).unwrap_or(0.0));
        constraints.push((grid_cell, minimum, maximum, preferred, percentage));
    }
    constraints.sort_by_key(|(cell, _, _, _, _)| (cell.col_span, cell.column_index));
    for (grid_cell, minimum, _, _, _) in &constraints {
        add_span_constraint(
            &mut minimums,
            grid_cell.column_index,
            grid_cell.col_span,
            *minimum,
        );
    }
    let mut maximums = minimums.clone();
    let mut preferences = minimums.clone();
    let mut priced = vec![false; col_count];
    for (grid_cell, _, maximum, preferred, _) in &constraints {
        add_span_constraint(
            &mut maximums,
            grid_cell.column_index,
            grid_cell.col_span,
            *maximum,
        );
        if let Some(preferred) = preferred {
            add_span_constraint(
                &mut preferences,
                grid_cell.column_index,
                grid_cell.col_span,
                *preferred,
            );
            let end = col_count.min(grid_cell.column_index + grid_cell.col_span);
            if grid_cell.column_index < end {
                priced[grid_cell.column_index..end].fill(true);
            }
        }
    }
    for column in 0..col_count {
        if priced[column] {
            maximums[column] = preferences[column];
        }
        if minimums[column] <= 0.0 {
            minimums[column] = base[column].min(if maximums[column] > 0.0 {
                maximums[column]
            } else {
                base[column]
            });
        }
        maximums[column] = maximums[column].max(minimums[column]);
        if maximums[column] <= 0.0 {
            maximums[column] = base[column];
        }
    }
    if explicit_width_px.is_none() && constraints.iter().any(|(_, _, _, _, pct)| pct.is_some()) {
        let mut table_width: f64 = maximums.iter().sum();
        let mut percentages = vec![0.0; col_count];
        for (cell, _, _, _, percentage) in &constraints {
            let Some(percentage) = percentage else {
                continue;
            };
            let start = cell.column_index;
            let end = col_count.min(start + cell.col_span);
            if start >= end {
                continue;
            }
            let natural: f64 = maximums[start..end].iter().sum();
            table_width = table_width.max(natural / percentage);
            add_span_constraint(&mut percentages, start, cell.col_span, *percentage);
        }
        let percentage_total: f64 = percentages.iter().sum();
        let remaining_columns: Vec<usize> = (0..col_count)
            .filter(|column| percentages[*column] <= 0.0)
            .collect();
        if percentage_total < 1.0 {
            let natural: f64 = remaining_columns
                .iter()
                .map(|column| maximums[*column])
                .sum();
            table_width = table_width.max(natural / (1.0 - percentage_total));
        }
        table_width = table_width.min(content_width);
        let (_, resolved) = autofit_content_widths(
            table_block,
            content_width,
            col_count,
            Some(table_width),
            content_widths,
        );
        maximums = resolved;
        if !remaining_columns.is_empty() {
            let total: f64 = maximums.iter().sum();
            let share = (table_width - total).max(0.0) / remaining_columns.len() as f64;
            for column in remaining_columns {
                maximums[column] += share;
            }
        } else if percentage_total < 1.0
            && minimums
                .iter()
                .zip(&percentages)
                .all(|(minimum, percentage)| *minimum / *percentage <= table_width)
        {
            let total: f64 = maximums.iter().sum();
            let share = (table_width - total).max(0.0) / percentage_total;
            for (width, percentage) in maximums.iter_mut().zip(&percentages) {
                *width += share * percentage;
            }
        }
    }
    (minimums, maximums)
}

/// Whether a cell holds content that shrinking would crop instead of rewrap:
/// an inline image, or any block other than a paragraph (a nested table, an
/// image, shape, chart or text box).
fn holds_rigid_content(blocks: &[crate::types::LayoutBlock]) -> bool {
    blocks.iter().any(|block| match block {
        crate::types::LayoutBlock::Paragraph(paragraph) => paragraph
            .runs
            .iter()
            .any(|run| matches!(run, crate::types::Run::Image(_))),
        _ => true,
    })
}

/// The widest inline image among a cell's paragraphs, which shrinking can't narrow.
fn widest_inline_image(blocks: &[crate::types::LayoutBlock]) -> f64 {
    blocks
        .iter()
        .filter_map(|block| match block {
            crate::types::LayoutBlock::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| &paragraph.runs)
        .filter_map(|run| match run {
            crate::types::Run::Image(image) => {
                Some(crate::measure_blocks::synthetic_inline_image_width(image))
            }
            _ => None,
        })
        .fold(0.0, f64::max)
}

/// A cell's narrowest width before shrinking crops or hides its content, and
/// whether that content is rigid (it can't rewrap at all). A rigid cell whose
/// minimum content width is unknown can't be narrowed safely at all.
fn cell_shrink_floor(
    table_block: &TableBlock,
    grid_cell: &ResolvedGridCell,
    content_widths: Option<&[Vec<Option<(f64, f64)>>]>,
) -> (f64, bool) {
    let cell = &table_block.rows[grid_cell.row_index].cells[grid_cell.cell_index];
    let padding = cell
        .padding
        .as_ref()
        .map_or(2.0 * DEFAULT_CELL_PADDING_X, |padding| {
            padding.left + padding.right
        });
    let floor = padding + widest_inline_image(&cell.blocks).max(1.0);
    let rigid = holds_rigid_content(&cell.blocks);
    if !rigid {
        return (floor, false);
    }
    let minimum = cell
        .min_content_width
        .or(content_widths
            .and_then(|rows| rows.get(grid_cell.row_index))
            .and_then(|cells| cells.get(grid_cell.cell_index))
            .copied()
            .flatten()
            .map(|widths| widths.0))
        .unwrap_or(f64::INFINITY);
    (floor.max(minimum), true)
}

fn span_width(widths: &[f64], grid_cell: &ResolvedGridCell) -> f64 {
    widths
        .iter()
        .skip(grid_cell.column_index)
        .take(grid_cell.col_span)
        .sum()
}

/// Autofit column widths, or `None` to keep the declared grid when they
/// would crop a cell's content.
fn resolve_autofit_column_widths(
    table_block: &TableBlock,
    content_width: f64,
    col_count: usize,
    explicit_width_px: Option<f64>,
    content_widths: Option<&[Vec<Option<(f64, f64)>>]>,
) -> Option<Vec<f64>> {
    let widths = autofit_column_widths(
        table_block,
        content_width,
        col_count,
        explicit_width_px,
        content_widths,
    )?;
    let crops = resolve_cell_grid(table_block).iter().any(|grid_cell| {
        let (floor, rigid) = cell_shrink_floor(table_block, grid_cell, content_widths);
        rigid && span_width(&widths, grid_cell) < floor
    });
    (!crops).then_some(widths)
}

fn autofit_column_widths(
    table_block: &TableBlock,
    content_width: f64,
    col_count: usize,
    explicit_width_px: Option<f64>,
    content_widths: Option<&[Vec<Option<(f64, f64)>>]>,
) -> Option<Vec<f64>> {
    let budget = table_width_budget(table_block, content_width);
    let (minimums, maximums) = autofit_content_widths(
        table_block,
        budget,
        col_count,
        explicit_width_px,
        content_widths,
    );
    let min_total: f64 = minimums.iter().sum();
    let max_total: f64 = maximums.iter().sum();
    let target = budget.min(
        min_total.max(explicit_width_px.unwrap_or(if max_total > 0.0 {
            max_total
        } else {
            budget
        })),
    );
    if target < min_total {
        let scale = target / min_total;
        let widths: Vec<f64> = minimums.into_iter().map(|width| width * scale).collect();
        let mut column_floors = vec![1.0_f64; col_count];
        let mut cells = resolve_cell_grid(table_block);
        cells.sort_by_key(|cell| (cell.col_span, cell.column_index));
        let mut below_cell_floor = false;
        for grid_cell in cells {
            let cell = &table_block.rows[grid_cell.row_index].cells[grid_cell.cell_index];
            let minimum = cell.min_content_width.or(content_widths
                .and_then(|rows| rows.get(grid_cell.row_index))
                .and_then(|cells| cells.get(grid_cell.cell_index))
                .copied()
                .flatten()
                .map(|widths| widths.0));
            if minimum.is_some_and(|minimum| span_width(&widths, &grid_cell) < minimum) {
                return None;
            }
            let (floor, _) = cell_shrink_floor(table_block, &grid_cell, content_widths);
            below_cell_floor |= span_width(&widths, &grid_cell) < floor;
            if floor.is_finite() {
                add_span_constraint(
                    &mut column_floors,
                    grid_cell.column_index,
                    grid_cell.col_span,
                    floor,
                );
            }
        }
        if below_cell_floor && content_width >= column_floors.iter().sum::<f64>() {
            return None;
        }
        return Some(widths);
    }
    if target >= max_total {
        return Some(distribute_to_target(maximums, target));
    }
    let flex: Vec<f64> = maximums
        .iter()
        .zip(&minimums)
        .map(|(max, min)| (max - min).max(0.0))
        .collect();
    let flex_total: f64 = flex.iter().sum();
    let extra = (target - min_total).max(0.0);
    if flex_total <= 0.0 {
        return Some(distribute_to_target(minimums, target));
    }
    Some(
        minimums
            .into_iter()
            .enumerate()
            .map(|(index, min)| min + extra * flex[index] / flex_total)
            .collect(),
    )
}

fn table_indent(table_block: &TableBlock) -> f64 {
    if matches!(
        table_block.justification.as_deref(),
        Some("center" | "right")
    ) {
        return 0.0;
    }
    table_block
        .indent
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
        .max(0.0)
}

/// The available width after the table's applied indent.
fn table_width_budget(table_block: &TableBlock, content_width: f64) -> f64 {
    (content_width - table_indent(table_block)).max(0.0)
}

/// Grid columns whose width no cell states, for a table that states no width
/// of its own and whose resolved `widths` leave room in `content_width`.
///
/// Word sizes exactly these columns from their content, so the declared
/// `w:gridCol` is only a hint and goes stale whenever the content changes.
/// Empty for fixed and autofit layouts, and whenever the declared
/// geometry already decides the answer.
pub fn content_sized_columns(
    table_block: &TableBlock,
    content_width: f64,
    widths: &[f64],
) -> Vec<usize> {
    if table_block.rows.is_empty() || widths.is_empty() {
        return Vec::new();
    }
    if table_block
        .width_algorithm
        .as_deref()
        .or(table_block.layout_mode.as_deref())
        .is_some_and(|algorithm| matches!(algorithm, "fixed" | "autofit"))
    {
        return Vec::new();
    }
    if preferred_width_px(
        table_block.preferred_width.as_ref(),
        table_block.width,
        table_block.width_type.as_deref(),
        content_width,
        None,
    )
    .is_some()
    {
        return Vec::new();
    }
    let total: f64 = widths.iter().sum();
    if !total.is_finite() || total >= table_width_budget(table_block, content_width) {
        return Vec::new();
    }
    let mut priced = vec![false; widths.len()];
    for grid_cell in resolve_cell_grid(table_block) {
        if grid_cell.col_span != 1 || grid_cell.column_index >= priced.len() {
            continue;
        }
        let Some(cell) = table_block
            .rows
            .get(grid_cell.row_index)
            .and_then(|row| row.cells.get(grid_cell.cell_index))
        else {
            continue;
        };
        if preferred_width_px(
            cell.preferred_width.as_ref(),
            cell.width_value,
            cell.width_type.as_deref(),
            total,
            cell.width,
        )
        .is_some()
        {
            priced[grid_cell.column_index] = true;
        }
    }
    (0..priced.len()).filter(|index| !priced[*index]).collect()
}

/// Raises each column toward `maximums[column]`, its widest unwrapped cell
/// content, spending only the room left inside the table's budget and sharing
/// that room in proportion to the demands when it cannot cover them all.
/// Columns never shrink, so a cell can only wrap onto fewer lines.
pub fn grow_content_sized_columns(
    table_block: &TableBlock,
    content_width: f64,
    maximums: &[f64],
    widths: &mut [f64],
) {
    let total: f64 = widths.iter().sum();
    let slack = table_width_budget(table_block, content_width) - total;
    if !(slack > 0.0) {
        return;
    }
    let demands: Vec<f64> = widths
        .iter()
        .enumerate()
        .map(|(index, width)| match maximums.get(index) {
            Some(maximum) if maximum.is_finite() => (maximum - width).max(0.0),
            _ => 0.0,
        })
        .collect();
    let demanded: f64 = demands.iter().sum();
    if !(demanded > 0.0) {
        return;
    }
    let share = (slack / demanded).min(1.0);
    for (width, demand) in widths.iter_mut().zip(&demands) {
        *width += demand * share;
    }
}

/// Resolves per-column pixel widths from the table's grid metadata and width
/// budget, per the module's three algorithms. Measures no cell content.
pub fn resolve_table_column_widths(table_block: &TableBlock, content_width: f64) -> Vec<f64> {
    resolve_table_column_widths_with_content(table_block, content_width, None)
}

pub(crate) fn resolve_table_column_widths_with_content(
    table_block: &TableBlock,
    content_width: f64,
    content_widths: Option<&[Vec<Option<(f64, f64)>>]>,
) -> Vec<f64> {
    let mut column_widths: Vec<f64> = table_block.column_widths.clone().unwrap_or_default();
    let explicit_width_px = preferred_width_px(
        table_block.preferred_width.as_ref(),
        table_block.width,
        table_block.width_type.as_deref(),
        content_width,
        None,
    );
    let col_count = count_table_columns(table_block);
    let target_width = explicit_width_px.unwrap_or(content_width);

    let algorithm = table_block
        .width_algorithm
        .as_deref()
        .or(table_block.layout_mode.as_deref())
        .unwrap_or("legacy");
    if !table_block.rows.is_empty() && algorithm == "fixed" {
        return resolve_fixed_column_widths(
            table_block,
            content_width,
            col_count,
            explicit_width_px,
        );
    }
    if !table_block.rows.is_empty() && algorithm == "autofit" {
        if let Some(widths) = resolve_autofit_column_widths(
            table_block,
            content_width,
            col_count,
            explicit_width_px,
            content_widths,
        ) {
            return widths;
        }
        if column_widths.is_empty() {
            column_widths = table_block.grid_widths.clone().unwrap_or_default();
        }
    }

    if !table_block.rows.is_empty() {
        column_widths = normalize_table_column_widths(&column_widths, col_count, target_width);
    }

    if !column_widths.is_empty()
        && let Some(explicit) = explicit_width_px
        && js_truthy(explicit)
    {
        let total: f64 = column_widths.iter().fold(0.0, |sum, &w| sum + w);
        if total > 0.0 && (total - explicit).abs() > 1.0 {
            let scale = explicit / total;
            column_widths = column_widths.into_iter().map(|w| w * scale).collect();
        }
    }

    column_widths
}

/// Total pixel width: the resolved columns, else the explicit table width,
/// else the whole content-width budget.
pub fn resolve_table_total_width_px(table_block: &TableBlock, content_width: f64) -> f64 {
    let column_widths = resolve_table_column_widths(table_block, content_width);
    let explicit_width_px = preferred_width_px(
        table_block.preferred_width.as_ref(),
        table_block.width,
        table_block.width_type.as_deref(),
        content_width,
        None,
    );
    let total = column_widths.iter().fold(0.0, |w, &cw| w + cw);
    if js_truthy(total) {
        return total;
    }
    if let Some(explicit) = explicit_width_px
        && js_truthy(explicit)
    {
        return explicit;
    }
    content_width
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};

    /// bun:test `toBeCloseTo(expected, digits)`: |actual - expected| < 0.5 * 10^-digits.
    fn assert_close_to(actual: f64, expected: f64, digits: i32) {
        assert!(
            (actual - expected).abs() < 0.5 * 10f64.powi(-digits),
            "expected {actual} to be close to {expected} ({digits} digits)"
        );
    }

    fn plain_cell() -> serde_json::Value {
        json!({ "id": 0, "blocks": [] })
    }

    fn table_with_column_widths(column_widths: Vec<f64>) -> TableBlock {
        let cells: Vec<serde_json::Value> = column_widths.iter().map(|_| plain_cell()).collect();
        serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": cells }],
            "columnWidths": column_widths,
        }))
        .unwrap()
    }

    #[test]
    fn dxa_twips_converted_to_pixels() {
        // 1440 twips = 1 inch = 96 px
        assert_close_to(
            resolve_table_width_px(Some(1440.0), Some("dxa"), 600.0).unwrap(),
            96.0,
            1,
        );
    }

    #[test]
    fn pct_fiftieths_of_a_percent_per_ecma_376() {
        assert_eq!(
            resolve_table_width_px(Some(2500.0), Some("pct"), 600.0),
            Some(300.0)
        );
        assert_eq!(
            resolve_table_width_px(Some(5000.0), Some("pct"), 600.0),
            Some(600.0)
        );
        // Small spec values must NOT be coerced to plain percent — `1` means 0.02%.
        assert_close_to(
            resolve_table_width_px(Some(1.0), Some("pct"), 5000.0).unwrap(),
            1.0,
            5,
        );
    }

    #[test]
    fn zero_negative_undefined_width_returns_none() {
        assert_eq!(resolve_table_width_px(Some(0.0), Some("dxa"), 600.0), None);
        assert_eq!(
            resolve_table_width_px(Some(-10.0), Some("dxa"), 600.0),
            None
        );
        assert_eq!(resolve_table_width_px(None, Some("dxa"), 600.0), None);
    }

    #[test]
    fn unrecognized_width_type_returns_none() {
        assert_eq!(
            resolve_table_width_px(Some(1440.0), Some("nil"), 600.0),
            None
        );
    }

    #[test]
    fn automatic_width_ignores_nonzero_values() {
        assert_eq!(
            resolve_table_width_px(Some(4500.0), Some("auto"), 600.0),
            None
        );
    }

    #[test]
    fn total_width_sums_explicit_column_widths() {
        assert_eq!(
            resolve_table_total_width_px(&table_with_column_widths(vec![200.0, 300.0]), 800.0),
            500.0
        );
    }

    #[test]
    fn total_width_falls_back_to_content_width_for_an_empty_table() {
        let empty: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [],
            "columnWidths": [],
        }))
        .unwrap();
        assert_eq!(resolve_table_total_width_px(&empty, 640.0), 640.0);
    }

    #[test]
    fn resolves_grid_positions_for_vertically_merged_cells() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [
                { "id": 0, "cells": [{ "id": 0, "blocks": [], "rowSpan": 3 }, plain_cell()] },
                { "id": 1, "cells": [plain_cell()] },
                { "id": 2, "cells": [plain_cell()] },
            ],
            "columnWidths": [100.0, 100.0],
        }))
        .unwrap();
        let g = |row_index, cell_index, column_index, col_span, row_span| ResolvedGridCell {
            row_index,
            cell_index,
            column_index,
            col_span,
            row_span,
        };
        assert_eq!(
            resolve_cell_grid(&block),
            vec![
                g(0, 0, 0, 1, 3),
                g(0, 1, 1, 1, 1),
                g(1, 0, 1, 1, 1),
                g(2, 0, 1, 1, 1),
            ]
        );
    }

    #[test]
    fn sparse_rows_and_explicit_grid_starts_define_the_grid() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{
                "id": 0,
                "gridBefore": 2,
                "gridAfter": 1,
                "cells": [{ "id": 0, "blocks": [], "gridStart": 3, "colSpan": 2 }],
            }],
        }))
        .unwrap();
        let resolved = resolve_cell_grid(&block);
        assert_eq!(resolved[0].column_index, 3);
        assert_eq!(count_table_columns(&block), 6);
    }

    #[test]
    fn fixed_layout_honors_first_row_cell_preferred_width_without_uniform_scaling() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "blocks": [], "preferredWidth": { "value": 3000, "type": "dxa" } },
                { "id": 1, "blocks": [] }
            ] }],
            "gridWidths": [100, 100],
            "layoutMode": "fixed",
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![200.0, 100.0]
        );
    }

    #[test]
    fn fixed_layout_replaces_first_row_grid_widths_before_reducing_the_total() {
        for (table_width, expected) in [
            (9000.0, [100.0, 500.0]),
            (6000.0, [200.0 / 3.0, 1000.0 / 3.0]),
        ] {
            let mut block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "fixed",
                "gridWidths": [twips_to_pixels(4500.0), twips_to_pixels(4500.0)],
                "preferredWidth": {"value": table_width, "type": "dxa"},
                "rows": [
                    {"id": 0, "cells": [
                        {"id": 0, "blocks": [], "preferredWidth": {"value": 1500, "type": "dxa"}},
                        {"id": 1, "blocks": [], "preferredWidth": {"value": 7500, "type": "dxa"}}
                    ]},
                    {"id": 1, "cells": [
                        {"id": 2, "blocks": [], "widthValue": 7500, "widthType": "dxa"},
                        {"id": 3, "blocks": [], "widthValue": 1500, "widthType": "dxa"}
                    ]}
                ]
            }))
            .unwrap();
            for legacy in [false, true] {
                if legacy {
                    for cell in &mut block.rows[0].cells {
                        let preferred = cell.preferred_width.take().unwrap();
                        cell.width_value = preferred.value;
                        cell.width_type = preferred.r#type;
                    }
                }
                let widths = resolve_table_column_widths(&block, 601.333_333);
                for (actual, expected) in widths.iter().zip(expected) {
                    assert_close_to(*actual, expected, 6);
                }
            }
        }
    }

    #[test]
    fn fixed_layout_enlarges_automatic_columns_in_proportion_to_the_grid() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "fixed",
            "gridWidths": [twips_to_pixels(1500.0), twips_to_pixels(4500.0)],
            "preferredWidth": {"value": 9000, "type": "dxa"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "preferredWidth": {"value": 0, "type": "auto"}},
                {"id": 1, "blocks": [], "preferredWidth": {"value": 0, "type": "auto"}}
            ]}]
        }))
        .unwrap();
        for legacy in [false, true] {
            if legacy {
                for cell in &mut block.rows[0].cells {
                    let preferred = cell.preferred_width.take().unwrap();
                    cell.width_value = preferred.value;
                    cell.width_type = preferred.r#type;
                }
            }
            assert_eq!(
                resolve_table_column_widths(&block, 601.333_333),
                vec![150.0, 450.0]
            );
        }
    }

    #[test]
    fn fixed_layout_replaces_spanning_first_row_preferences() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "fixed", "gridWidths": [300, 300, 300],
            "preferredWidth": {"value": 9000, "type": "dxa"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "colSpan": 2,
                 "preferredWidth": {"value": 1500, "type": "dxa"}},
                {"id": 1, "blocks": [], "preferredWidth": {"value": 7500, "type": "dxa"}}
            ]}]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![50.0, 50.0, 500.0]
        );
    }

    #[test]
    fn autofit_uses_intrinsic_min_and_max_widths_and_shrinks_to_content() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "blocks": [], "minContentWidth": 50, "maxContentWidth": 150 },
                { "id": 1, "blocks": [], "minContentWidth": 100, "maxContentWidth": 200 }
            ] }],
            "gridWidths": [300, 300],
            "layoutMode": "autofit",
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![150.0, 200.0]
        );
    }

    #[test]
    fn autofit_whose_indent_leaves_too_little_room_keeps_its_grid() {
        for (indent, grid) in [
            (600.0, json!({"columnWidths": [100, 100]})),
            (599.0, json!({"columnWidths": [100, 100]})),
            (600.0, json!({"gridWidths": [100, 100]})),
            (599.0, json!({"gridWidths": [100, 100]})),
        ] {
            let mut table = json!({
                "id": 0,
                "rows": [{ "id": 0, "cells": [
                    { "id": 0, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40 },
                    { "id": 1, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40 }
                ] }],
                "indent": indent,
                "layoutMode": "autofit",
            });
            table
                .as_object_mut()
                .unwrap()
                .extend(grid.as_object().unwrap().clone());
            let block: TableBlock = serde_json::from_value(table).unwrap();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![100.0, 100.0],
                "{indent}"
            );
        }
    }

    #[test]
    fn autofit_with_zero_padding_and_exhausted_indent_keeps_its_grid() {
        for grid in [
            json!({"columnWidths": [100, 100]}),
            json!({"gridWidths": [100, 100]}),
        ] {
            let mut table = json!({
                "id": 0, "layoutMode": "autofit", "indent": 600,
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40,
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                    {"id": 1, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40,
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
                ]}]
            });
            table
                .as_object_mut()
                .unwrap()
                .extend(grid.as_object().unwrap().clone());
            let block: TableBlock = serde_json::from_value(table).unwrap();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![100.0, 100.0]
            );
        }
    }

    #[test]
    fn autofit_keeps_its_grid_when_each_column_would_shrink_below_its_minimum() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [100, 100], "indent": 500,
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300},
                {"id": 1, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300}
            ]}]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        block.indent = Some(570.0);
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        block.indent = Some(571.0);
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        block.indent = Some(19.0);
        assert_eq!(
            resolve_table_column_widths(&block, 20.0),
            vec![100.0, 100.0]
        );
    }

    #[test]
    fn autofit_keeps_its_grid_with_known_minimums_and_varying_padding() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [100, 100], "indent": 500,
            "rows": [
                {"id": 0, "cells": [
                    {"id": 0, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300,
                     "padding": {"top": 0, "bottom": 0, "left": 40, "right": 20}},
                    {"id": 1, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300,
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
                ]},
                {"id": 1, "cells": [
                    {"id": 2, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300,
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                    {"id": 3, "blocks": [], "minContentWidth": 300, "maxContentWidth": 300,
                     "padding": {"top": 0, "bottom": 0, "left": 30, "right": 30}}
                ]}
            ]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        let padding = block.rows[1].cells[1].padding.as_mut().unwrap();
        padding.left = 20.0;
        padding.right = 20.0;
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        let padding = block.rows[0].cells[0].padding.as_mut().unwrap();
        padding.left = 29.0;
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        block.indent = Some(599.0);
        for row in &mut block.rows {
            for cell in &mut row.cells {
                let padding = cell.padding.as_mut().unwrap();
                padding.left = 0.0;
                padding.right = 0.0;
            }
        }
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
        block.indent = Some(598.0);
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![100.0, 100.0]
        );
    }

    #[test]
    fn autofit_keeps_its_grid_when_shrinking_would_crop_an_inline_image() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [200, 400],
            "preferredWidth": {"value": 9000, "type": "dxa"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "minContentWidth": 200, "maxContentWidth": 200,
                 "blocks": [{"kind": "paragraph", "id": 0, "runs": [
                     {"kind": "image", "src": "", "width": 200, "height": 40}
                 ]}],
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                {"id": 1, "blocks": [], "minContentWidth": 1510.16, "maxContentWidth": 1510.16,
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
            ]}]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![200.0, 400.0]
        );
    }

    #[test]
    fn autofit_keeps_its_grid_when_shrinking_would_crop_an_indented_image_or_nested_table() {
        let image = json!({"kind": "paragraph", "id": 0,
            "attrs": {"indent": {"left": 100}},
            "runs": [{"kind": "image", "src": "", "width": 200, "height": 40}]});
        let nested = json!({"kind": "table", "id": 1, "layoutMode": "fixed", "columnWidths": [200],
        "rows": [{"id": 0, "cells": [{"id": 0, "blocks": [
            {"kind": "paragraph", "id": 2, "runs": [
                {"kind": "image", "src": "", "width": 200, "height": 40}
            ]}
        ]}]}]});
        let sibling = json!([{"kind": "paragraph", "id": 3, "runs": [
            {"kind": "image", "src": "", "width": 1, "height": 1}
        ]}]);
        for (grid, content, minimum, other, other_blocks) in [
            ([300.0, 300.0], image.clone(), 300.0, 498.35, json!([])),
            ([200.0, 400.0], nested, 200.0, 1510.16, json!([])),
            ([300.0, 300.0], image, 300.0, 498.35, sibling),
        ] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": grid,
                "preferredWidth": {"value": 9000, "type": "dxa"},
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [content],
                     "minContentWidth": minimum, "maxContentWidth": minimum,
                     "preferredWidth": {"value": 0, "type": "auto"},
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                    {"id": 1, "blocks": other_blocks,
                     "minContentWidth": other, "maxContentWidth": other,
                     "preferredWidth": {"value": 0, "type": "auto"},
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
                ]}]
            }))
            .unwrap();
            assert_eq!(resolve_table_column_widths(&block, 600.0), grid.to_vec());
        }
    }

    #[test]
    fn autofit_keeps_its_grid_when_a_preferred_width_would_crop_an_unmeasured_image() {
        for (indent, preferred) in [(0, 1500), (100, 3750)] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [{"kind": "paragraph", "id": 0,
                        "attrs": {"indent": {"left": indent}},
                        "runs": [
                            {"kind": "image", "src": "", "width": 200, "height": 40},
                            {"kind": "text", "text": "x"}
                        ]}],
                     "preferredWidth": {"value": preferred, "type": "dxa"},
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                    {"id": 1, "blocks": [],
                     "preferredWidth": {"value": 0, "type": "auto"},
                     "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
                ]}]
            }))
            .unwrap();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![300.0, 300.0],
                "indent {indent}"
            );
        }
    }

    #[test]
    fn autofit_keeps_its_grid_when_shrinking_would_clip_known_text_minimums() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "preferredWidth": {"value": 9000, "type": "dxa"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [{"kind": "paragraph", "id": 0, "runs": [
                    {"kind": "text", "text": "W"}
                ]}],
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                {"id": 1, "blocks": [{"kind": "paragraph", "id": 1, "runs": [
                    {"kind": "text", "text": "W".repeat(100)}
                ]}],
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
            ]}]
        }))
        .unwrap();
        let content = vec![vec![Some((15.1016, 15.1016)), Some((1510.16, 1510.16))]];
        assert_eq!(
            resolve_table_column_widths_with_content(&block, 600.0, Some(&content)),
            vec![300.0, 300.0]
        );
    }

    #[test]
    fn autofit_nested_tables_keep_the_outer_grid_and_child_grid_fallback() {
        use crate::measure_blocks::{MeasurementConfig, measure_block};
        use crate::types::{BlockExtent, LayoutBlock};

        let mut outer: LayoutBlock = serde_json::from_value(json!({
            "kind": "table", "id": 0, "layoutMode": "autofit", "gridWidths": [300],
            "preferredWidth": {"value": 0, "type": "auto"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
                 "blocks": [{"kind": "table", "id": 1, "layoutMode": "autofit", "gridWidths": [300],
                     "preferredWidth": {"value": 0, "type": "auto"},
                     "rows": [{"id": 1, "cells": [
                         {"id": 1, "preferredWidth": {"value": 1500, "type": "dxa"},
                          "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
                          "blocks": [{"kind": "paragraph", "id": 0, "runs": [
                              {"kind": "image", "src": "", "width": 200, "height": 40},
                              {"kind": "text", "text": "x"}
                          ]}]}
                     ]}]
                 }]}
            ]}]
        }))
        .unwrap();
        let LayoutBlock::Table(table) = &outer else {
            panic!()
        };
        let LayoutBlock::Table(child) = &table.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        let LayoutBlock::Paragraph(paragraph) = &child.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        let config = MeasurementConfig {
            defaults: json!({"fontFamily": "Arial", "fontSize": 12}),
            ..MeasurementConfig::default()
        };
        let content = vec![vec![crate::typed_measure::intrinsic_widths(
            paragraph, 600.0, &config,
        )]];
        assert_eq!(content, vec![vec![None]]);
        assert_eq!(
            resolve_table_column_widths_with_content(child, 600.0, Some(&content)),
            vec![300.0]
        );
        let BlockExtent::Table(measured) = measure_block(&mut outer, 600.0, &config).unwrap()
        else {
            panic!()
        };
        assert_eq!(measured.column_widths, vec![300.0]);
        let cell = &measured.rows[0].cells[0];
        assert_eq!(cell.width, 300.0);
        let BlockExtent::Table(child) = &cell.blocks[0] else {
            panic!()
        };
        assert_eq!(child.column_widths.iter().sum::<f64>(), 300.0);
        assert!(child.column_widths.iter().sum::<f64>() <= cell.width);
    }

    #[test]
    fn autofit_keeps_its_grid_when_shrinking_would_hide_another_cells_text() {
        for padding in [
            Value::Null,
            json!({"top": 0, "bottom": 0, "left": 7, "right": 7}),
        ] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
                "preferredWidth": {"value": 9000, "type": "dxa"},
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "minContentWidth": 4000, "maxContentWidth": 4000,
                     "preferredWidth": {"value": 0, "type": "auto"}, "padding": padding},
                    {"id": 1, "blocks": [], "minContentWidth": 49, "maxContentWidth": 49,
                     "preferredWidth": {"value": 0, "type": "auto"}, "padding": padding}
                ]}]
            }))
            .unwrap();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![300.0, 300.0]
            );
        }
    }

    #[test]
    fn autofit_keeps_its_grid_when_shrinking_below_merged_cell_minimums() {
        for merged_first in [true, false] {
            let merged = json!({"id": 0, "cells": [
                {"id": 0, "blocks": [], "colSpan": 2,
                 "minContentWidth": 40, "maxContentWidth": 40,
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 7, "right": 7}}
            ]});
            let individual = json!({"id": 1, "cells": [
                {"id": 1, "blocks": [], "minContentWidth": 20, "maxContentWidth": 20,
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}},
                {"id": 2, "blocks": [], "minContentWidth": 1180, "maxContentWidth": 1180,
                 "preferredWidth": {"value": 0, "type": "auto"},
                 "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0}}
            ]});
            let rows = if merged_first {
                vec![merged, individual]
            } else {
                vec![individual, merged]
            };
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300], "rows": rows
            }))
            .unwrap();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![300.0, 300.0]
            );
        }
    }

    #[test]
    fn table_width_budget_matches_placement_alignment() {
        for justification in [
            None,
            Some("left"),
            Some("start"),
            Some("end"),
            Some("center"),
            Some("right"),
        ] {
            for bidi in [false, true] {
                let mut block: TableBlock = serde_json::from_value(json!({
                    "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
                    "preferredWidth": {"value": 9000, "type": "dxa"},
                    "indent": 100, "justification": justification, "bidi": bidi,
                    "rows": [{"id": 0, "cells": [
                        {"id": 0, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40},
                        {"id": 1, "blocks": [], "minContentWidth": 40, "maxContentWidth": 40}
                    ]}]
                }))
                .unwrap();
                let indent = if matches!(justification, Some("center" | "right")) {
                    0.0
                } else {
                    100.0
                };
                for legacy in [false, true] {
                    if legacy {
                        let preferred = block.preferred_width.take().unwrap();
                        block.width = preferred.value;
                        block.width_type = preferred.r#type;
                    }
                    assert_eq!(
                        resolve_table_column_widths(&block, 600.0),
                        vec![(600.0 - indent) / 2.0; 2],
                        "{justification:?}, bidi={bidi}, legacy={legacy}"
                    );
                }
            }
        }
    }

    #[test]
    fn autofit_cell_percentages_use_the_content_sized_table_width() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "rows": [{"id": 0, "cells": [
                {"id": 0, "preferredWidth": {"value": 2500, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 0, "runs": [{"kind": "text", "text": "Hello"}]}]},
                {"id": 1, "preferredWidth": {"value": 2500, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 1, "runs": [{"kind": "text", "text": "Hello"}]}]}
            ]}]
        }))
        .unwrap();
        let content = vec![vec![Some((36.0, 36.0)), Some((36.0, 36.0))]];
        for legacy in [false, true] {
            if legacy {
                for cell in &mut block.rows[0].cells {
                    let preferred = cell.preferred_width.take().unwrap();
                    cell.width_value = preferred.value;
                    cell.width_type = preferred.r#type;
                }
            }
            let widths =
                resolve_table_column_widths_with_content(&block, 601.333_333, Some(&content));
            for width in widths {
                assert_close_to(width, 36.0, 6);
            }
        }
        block.width = Some(9000.0);
        block.width_type = Some("dxa".to_owned());
        assert_eq!(
            resolve_table_column_widths_with_content(&block, 601.333_333, Some(&content)),
            vec![300.0, 300.0]
        );
    }

    #[test]
    fn autofit_percentages_balance_unequal_content_and_unpriced_columns() {
        for (first, second, expected) in [
            (Some(2500), Some(2500), [72.0, 72.0]),
            (Some(2500), None, [72.0, 72.0]),
            (Some(1250), Some(3750), [36.0, 108.0]),
            (Some(3750), None, [216.0, 72.0]),
        ] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "widthValue": first, "widthType": "pct",
                     "minContentWidth": 36, "maxContentWidth": 36},
                    {"id": 1, "blocks": [], "widthValue": second, "widthType": "pct",
                     "minContentWidth": 72, "maxContentWidth": 72}
                ]}]
            }))
            .unwrap();
            let widths = resolve_table_column_widths(&block, 601.333_333);
            for (actual, expected) in widths.iter().zip(expected) {
                assert_close_to(*actual, expected, 6);
            }
        }
    }

    #[test]
    fn autofit_preserves_content_sized_width_when_all_percentages_total_less_than_100() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "rows": [{"id": 0, "cells": [
                {"id": 0, "preferredWidth": {"value": 1250, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 0, "runs": [{"kind": "text", "text": "Hello"}]}]},
                {"id": 1, "preferredWidth": {"value": 1250, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 1, "runs": [{"kind": "text", "text": "HelloHello"}]}]}
            ]}]
        }))
        .unwrap();
        let content = vec![vec![Some((36.4675, 36.4675)), Some((72.935, 72.935))]];
        for legacy in [false, true] {
            if legacy {
                for cell in &mut block.rows[0].cells {
                    let preferred = cell.preferred_width.take().unwrap();
                    cell.width_value = preferred.value;
                    cell.width_type = preferred.r#type;
                }
            }
            let widths =
                resolve_table_column_widths_with_content(&block, 601.333_333, Some(&content));
            assert_eq!(widths.len(), 2);
            for width in widths {
                assert_close_to(width, 145.87, 6);
            }
        }
    }

    #[test]
    fn autofit_preserves_unequal_ratios_when_all_percentages_total_less_than_100() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "rows": [{"id": 0, "cells": [
                {"id": 0, "preferredWidth": {"value": 1250, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 0, "runs": [{"kind": "text", "text": "Hello"}]}]},
                {"id": 1, "preferredWidth": {"value": 2500, "type": "pct"},
                 "blocks": [{"kind": "paragraph", "id": 1, "runs": [{"kind": "text", "text": "HelloHello"}]}]}
            ]}]
        }))
        .unwrap();
        let content = vec![vec![Some((36.4675, 36.4675)), Some((72.935, 72.935))]];
        for legacy in [false, true] {
            if legacy {
                for cell in &mut block.rows[0].cells {
                    let preferred = cell.preferred_width.take().unwrap();
                    cell.width_value = preferred.value;
                    cell.width_type = preferred.r#type;
                }
            }
            let widths =
                resolve_table_column_widths_with_content(&block, 601.333_333, Some(&content));
            assert_eq!(widths.len(), 2);
            for (actual, expected) in widths.iter().zip([48.60, 97.27]) {
                assert_close_to(*actual, expected, 1);
            }
            assert_close_to(widths[1] / widths[0], 2.0, 6);
            assert_close_to(widths.iter().sum(), 145.87, 6);
        }
    }

    #[test]
    fn autofit_automatic_cell_widths_ignore_numeric_and_pixel_preferences() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "width": 9000, "widthType": "auto",
            "rows": [{"id": 0, "cells": [
                {"id": 0, "width": 300, "preferredWidth": {"value": 4500, "type": "auto"},
                 "minContentWidth": 36, "maxContentWidth": 36,
                 "blocks": [{"kind": "paragraph", "id": 0, "runs": [{"kind": "text", "text": "Hello"}]}]},
                {"id": 1, "width": 300, "preferredWidth": {"value": 4500, "type": "auto"},
                 "minContentWidth": 36, "maxContentWidth": 36,
                 "blocks": [{"kind": "paragraph", "id": 1, "runs": [{"kind": "text", "text": "Hello"}]}]}
            ]}]
        }))
        .unwrap();
        for legacy in [false, true] {
            if legacy {
                for cell in &mut block.rows[0].cells {
                    let preferred = cell.preferred_width.take().unwrap();
                    cell.width_value = preferred.value;
                    cell.width_type = preferred.r#type;
                }
            }
            let widths = resolve_table_column_widths(&block, 601.333_333);
            for width in widths {
                assert_close_to(width, 36.0, 6);
            }
        }
    }

    #[test]
    fn fixed_layout_scales_oversubscribed_grids_and_preferences_to_table_width() {
        for (grid, preferred, expected) in [
            ([300.0, 300.0], [200.0, 200.0], [200.0, 200.0]),
            ([100.0, 100.0], [450.0, 150.0], [300.0, 100.0]),
        ] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "fixed", "gridWidths": grid,
                "preferredWidth": {"value": 6000, "type": "dxa"},
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "width": preferred[0]},
                    {"id": 1, "blocks": [], "width": preferred[1]}
                ]}]
            }))
            .unwrap();
            let widths = resolve_table_column_widths(&block, 600.0);
            for (actual, expected) in widths.iter().zip(expected) {
                assert_close_to(*actual, expected, 6);
            }
        }
    }

    #[test]
    fn autofit_grid_fallback_preserves_unbreakable_minimums_and_explicit_table_widths() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [600],
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "minContentWidth": 1510.16, "maxContentWidth": 1510.16}
            ]}]
        }))
        .unwrap();
        assert_close_to(resolve_table_column_widths(&block, 600.0)[0], 600.0, 6);
        assert_eq!(resolve_table_column_widths(&block, 0.0), vec![600.0]);

        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
            "preferredWidth": {"value": 15000, "type": "dxa"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "minContentWidth": 900, "maxContentWidth": 1000},
                {"id": 1, "blocks": [], "minContentWidth": 300, "maxContentWidth": 500}
            ]}]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![500.0, 500.0]
        );
        block.indent = Some(60.0);
        assert_eq!(
            resolve_table_column_widths(&block, 600.0),
            vec![500.0, 500.0]
        );
    }

    #[test]
    fn autofit_grid_fallback_preserves_percentage_table_width_with_indent() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300], "indent": 60,
            "preferredWidth": {"value": 10000, "type": "pct"},
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "minContentWidth": 900, "maxContentWidth": 1000},
                {"id": 1, "blocks": [], "minContentWidth": 300, "maxContentWidth": 500}
            ]}]
        }))
        .unwrap();
        for legacy in [false, true] {
            if legacy {
                let preferred = block.preferred_width.take().unwrap();
                block.width = preferred.value;
                block.width_type = preferred.r#type;
            }
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![600.0, 600.0]
            );
        }
    }

    #[test]
    fn autofit_exceeds_preferred_cell_width_only_for_content_minimums() {
        for (minimum, maximum, expected) in [
            (40.0, 238.34, 100.0),
            (120.0, 238.34, 120.0),
            (20.0, 50.0, 100.0),
        ] {
            let block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300],
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "preferredWidth": {"value": 1500, "type": "dxa"},
                     "minContentWidth": minimum, "maxContentWidth": maximum}
                ]}]
            }))
            .unwrap();
            assert_eq!(resolve_table_column_widths(&block, 600.0), vec![expected]);
        }
    }

    #[test]
    fn autofit_caps_unpriced_cells_in_a_column_with_a_preferred_width() {
        let mut block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [300],
            "rows": [
                {"id": 0, "cells": [
                    {"id": 0, "blocks": [], "preferredWidth": {"value": 1500, "type": "dxa"},
                     "minContentWidth": 20, "maxContentWidth": 50}
                ]},
                {"id": 1, "cells": [
                    {"id": 1, "blocks": [], "minContentWidth": 40, "maxContentWidth": 238.34}
                ]}
            ]
        }))
        .unwrap();
        assert_eq!(resolve_table_column_widths(&block, 600.0), vec![100.0]);
        block.rows[1].cells[0].min_content_width = Some(120.0);
        assert_eq!(resolve_table_column_widths(&block, 600.0), vec![120.0]);
    }

    #[test]
    fn autofit_applies_individual_constraints_before_spanning_cells_in_any_row_order() {
        for spanning_preferred in [false, true] {
            let mut block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [300, 300],
                "rows": [
                    {"id": 0, "cells": [
                        {"id": 0, "blocks": [], "colSpan": 2,
                         "minContentWidth": 300, "maxContentWidth": 300}
                    ]},
                    {"id": 1, "cells": [
                        {"id": 1, "blocks": [], "minContentWidth": 20, "maxContentWidth": 20},
                        {"id": 2, "blocks": [], "minContentWidth": 280, "maxContentWidth": 280}
                    ]}
                ]
            }))
            .unwrap();
            if spanning_preferred {
                block.rows[0].cells[0].width = Some(300.0);
            }
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![20.0, 280.0]
            );
            block.rows.reverse();
            assert_eq!(
                resolve_table_column_widths(&block, 600.0),
                vec![20.0, 280.0]
            );
        }
    }

    #[test]
    fn autofit_nowrap_protects_absolute_preferences_but_uses_natural_width_for_auto_and_pct() {
        for (preferred, expected) in [
            (json!({"value": 1500, "type": "dxa"}), 100.0),
            (json!({"value": 0, "type": "auto"}), 642.99),
            (json!({"value": 4500, "type": "auto"}), 642.99),
            (json!({"value": 500, "type": "pct"}), 642.99),
        ] {
            let mut block: TableBlock = serde_json::from_value(json!({
                "id": 0, "layoutMode": "autofit", "gridWidths": [100],
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [], "preferredWidth": preferred,
                     "minContentWidth": 40, "maxContentWidth": 642.99, "noWrap": true}
                ]}]
            }))
            .unwrap();
            assert_eq!(resolve_table_column_widths(&block, 1000.0), vec![expected]);
            let cell = &mut block.rows[0].cells[0];
            let preferred = cell.preferred_width.take().unwrap();
            cell.width_value = preferred.value;
            cell.width_type = preferred.r#type;
            assert_eq!(resolve_table_column_widths(&block, 1000.0), vec![expected]);
        }
    }

    #[test]
    fn autofit_nowrap_absolute_preferences_are_floors_during_flexible_shrinkage() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0, "layoutMode": "autofit", "gridWidths": [100, 200],
            "rows": [{"id": 0, "cells": [
                {"id": 0, "blocks": [], "widthValue": 1500, "widthType": "dxa",
                 "minContentWidth": 40, "maxContentWidth": 642.99, "noWrap": true},
                {"id": 1, "blocks": [], "minContentWidth": 20, "maxContentWidth": 200}
            ]}]
        }))
        .unwrap();
        assert_eq!(
            resolve_table_column_widths(&block, 150.0),
            vec![100.0, 50.0]
        );
    }

    /// `oxi-en-administrative-04`, measured off Word's own `reference.pdf`: an
    /// `auto` first column whose `w:gridCol` of 2143tw no longer fits the
    /// heading Word lays out at 110.028pt, so Word widens it to 111.805pt and
    /// leaves the three priced columns on their `w:tcW`.
    fn administrative_04_table() -> TableBlock {
        let priced =
            |value: f64| json!({ "id": 0, "blocks": [], "widthValue": value, "widthType": "dxa" });
        serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "blocks": [], "widthValue": 0, "widthType": "auto" },
                priced(1821.0),
                priced(2410.0),
                priced(2410.0),
            ] }],
            "columnWidths": [142.866_666, 121.4, 160.666_666, 160.666_666],
            "width": 0,
            "widthType": "auto",
        }))
        .unwrap()
    }

    #[test]
    fn an_unpriced_column_widens_to_its_content_inside_the_leftover_budget() {
        let block = administrative_04_table();
        let mut widths = resolve_table_column_widths(&block, 601.333_333);
        assert_eq!(content_sized_columns(&block, 601.333_333, &widths), vec![0]);
        // 110.028pt of heading plus the 15tw cell margins Word reserves.
        grow_content_sized_columns(&block, 601.333_333, &[148.704, 0.0, 0.0, 0.0], &mut widths);
        assert_close_to(widths[0], 148.704, 3);
        assert_close_to(widths[1], 121.4, 3);
        assert_close_to(widths[2], 160.666_666, 3);
        assert_close_to(widths[3], 160.666_666, 3);
        // Word's own rules sit 0.37px further out; the declared grid was 5.84px short.
        assert!((widths[0] - 149.073).abs() < 0.5);
    }

    #[test]
    fn a_table_that_states_its_own_width_keeps_the_declared_grid() {
        let mut block = administrative_04_table();
        block.width = Some(8784.0);
        block.width_type = Some("dxa".to_owned());
        let widths = resolve_table_column_widths(&block, 601.333_333);
        assert_eq!(
            content_sized_columns(&block, 601.333_333, &widths),
            Vec::<usize>::new()
        );
    }

    #[test]
    fn demands_beyond_the_leftover_budget_are_shared_in_proportion() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [plain_cell(), plain_cell()] }],
            "columnWidths": [100.0, 100.0],
        }))
        .unwrap();
        let mut widths = resolve_table_column_widths(&block, 260.0);
        assert_eq!(content_sized_columns(&block, 260.0, &widths), vec![0, 1]);
        grow_content_sized_columns(&block, 260.0, &[160.0, 120.0], &mut widths);
        assert_close_to(widths[0], 145.0, 6);
        assert_close_to(widths[1], 115.0, 6);
        assert_close_to(widths[0] + widths[1], 260.0, 6);
    }

    #[test]
    fn a_fixed_layout_table_is_never_content_sized() {
        let mut block = administrative_04_table();
        block.layout_mode = Some("fixed".to_owned());
        let widths = resolve_table_column_widths(&block, 601.333_333);
        assert_eq!(
            content_sized_columns(&block, 601.333_333, &widths),
            Vec::<usize>::new()
        );
    }

    #[test]
    fn a_grid_that_already_fills_the_budget_never_grows() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [plain_cell(), plain_cell()] }],
            "columnWidths": [100.0, 100.0],
        }))
        .unwrap();
        let widths = resolve_table_column_widths(&block, 200.0);
        assert_eq!(
            content_sized_columns(&block, 200.0, &widths),
            Vec::<usize>::new()
        );
    }
}
