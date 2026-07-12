use std::collections::{HashMap, HashSet};

use serde::Serialize;

use crate::types::TableBlock;

const TWIPS_PER_INCH: f64 = 1440.0;
/// Pixels per inch at the standard 96 DPI assumption.
const PIXELS_PER_INCH: f64 = 96.0;

fn twips_to_pixels(twips: f64) -> f64 {
    (twips / TWIPS_PER_INCH) * PIXELS_PER_INCH
}

fn is_nonzero_number(v: f64) -> bool {
    v != 0.0 && !v.is_nan()
}

/// Resolve a DOCX width pair to pixels. `pct` values are 50ths of a percent
/// (ECMA-376 §17.18.111 — 5000 means 100%). `dxa` / `auto` / unset are twips.
pub fn resolve_table_width_px(
    value: Option<f64>,
    width_type: Option<&str>,
    parent_width: f64,
) -> Option<f64> {
    let value = value?;
    if value <= 0.0 || value.is_nan() {
        return None;
    }
    if width_type == Some("pct") {
        return Some((parent_width * value) / 5000.0);
    }
    if width_type.is_none() || width_type == Some("dxa") || width_type == Some("auto") {
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

/// Resolve every cell's grid column index, accounting for `colSpan` and the
/// columns occupied by vertically-merged (`rowSpan`) cells from earlier rows.
///
/// Single source of truth for table grid geometry — width-free on purpose:
/// callers multiply `column_index` by their own (possibly scaled) column
/// widths to get an x offset.
pub fn resolve_cell_grid(table_block: &TableBlock) -> Vec<ResolvedGridCell> {
    let mut occupied: HashMap<usize, HashSet<usize>> = HashMap::new();
    let mut out: Vec<ResolvedGridCell> = Vec::new();
    for row_index in 0..table_block.rows.len() {
        let cells = &table_block.rows[row_index].cells;
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

/// Total grid columns, derived from the widest row's accumulated colSpans.
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

fn preferred_width_px(
    preferred: Option<&crate::types::PreferredWidth>,
    legacy_value: Option<f64>,
    legacy_type: Option<&str>,
    parent_width: f64,
    legacy_px: Option<f64>,
) -> Option<f64> {
    preferred
        .and_then(|width| {
            resolve_table_width_px(width.value, width.r#type.as_deref(), parent_width)
        })
        .or_else(|| resolve_table_width_px(legacy_value, legacy_type, parent_width))
        .or_else(|| legacy_px.filter(|value| *value > 0.0))
}

fn add_span_constraint(widths: &mut [f64], start: usize, span: usize, required: f64) {
    if required <= 0.0 || required.is_nan() || start >= widths.len() {
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
            explicit_width_px.unwrap_or(content_width),
            cell.width,
        ) {
            add_span_constraint(
                &mut widths,
                grid_cell.column_index,
                grid_cell.col_span,
                preferred,
            );
        }
    }
    explicit_width_px.map_or(widths.clone(), |target| {
        distribute_to_target(widths, target)
    })
}

fn resolve_autofit_column_widths(
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
    let base = normalize_table_column_widths(
        source,
        col_count,
        explicit_width_px.unwrap_or(content_width),
    );
    let mut minimums = vec![0.0; col_count];
    let mut maximums = vec![0.0; col_count];
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
            explicit_width_px.unwrap_or(content_width),
            cell.width,
        );
        let mut minimum = cell.min_content_width.unwrap_or(0.0).max(0.0);
        if cell.no_wrap.unwrap_or(false) {
            minimum = minimum.max(cell.max_content_width.unwrap_or(0.0));
        }
        let maximum = minimum.max(cell.max_content_width.or(preferred).unwrap_or(0.0));
        add_span_constraint(
            &mut minimums,
            grid_cell.column_index,
            grid_cell.col_span,
            minimum,
        );
        add_span_constraint(
            &mut maximums,
            grid_cell.column_index,
            grid_cell.col_span,
            maximum,
        );
        if let Some(preferred) = preferred {
            add_span_constraint(
                &mut maximums,
                grid_cell.column_index,
                grid_cell.col_span,
                preferred,
            );
        }
    }
    for column in 0..col_count {
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
    let min_total: f64 = minimums.iter().sum();
    let max_total: f64 = maximums.iter().sum();
    let target = min_total.max(content_width.min(explicit_width_px.unwrap_or(
        if max_total > 0.0 {
            max_total
        } else {
            content_width
        },
    )));
    if target >= max_total {
        return distribute_to_target(maximums, target);
    }
    let flex: Vec<f64> = maximums
        .iter()
        .zip(&minimums)
        .map(|(max, min)| (max - min).max(0.0))
        .collect();
    let flex_total: f64 = flex.iter().sum();
    let extra = (target - min_total).max(0.0);
    if flex_total <= 0.0 {
        return distribute_to_target(minimums, target);
    }
    minimums
        .into_iter()
        .enumerate()
        .map(|(index, min)| min + extra * flex[index] / flex_total)
        .collect()
}

/// Make `column_widths` exactly `col_count` long with every entry positive.
/// Missing trailing columns inherit the average of existing positives; zero
/// or negative entries split the leftover `target_width` evenly. Callers
/// scale down totals that exceed the target — this helper only fills gaps.
pub fn normalize_table_column_widths(
    column_widths: &[f64],
    col_count: usize,
    target_width: f64,
) -> Vec<f64> {
    if col_count == 0 {
        return Vec::new();
    }

    let even_width = if target_width > 0.0 {
        target_width / col_count as f64
    } else {
        0.0
    };

    if column_widths.is_empty() {
        return vec![even_width; col_count];
    }

    let mut normalized: Vec<f64> = column_widths.iter().copied().take(col_count).collect();
    let missing_columns = col_count - normalized.len();
    if missing_columns > 0 {
        let existing_positive: Vec<f64> = normalized.iter().copied().filter(|w| *w > 0.0).collect();
        let fallback_width = if !existing_positive.is_empty() {
            existing_positive.iter().fold(0.0, |sum, w| sum + w) / existing_positive.len() as f64
        } else {
            even_width
        };
        normalized.extend(std::iter::repeat_n(fallback_width, missing_columns));
    }

    let positive_total = normalized
        .iter()
        .fold(0.0, |sum, &w| sum + if w > 0.0 { w } else { 0.0 });
    let non_positive_count = normalized.iter().filter(|&&w| w <= 0.0).count();

    if positive_total <= 0.0 {
        return vec![even_width; col_count];
    }
    if non_positive_count == 0 {
        return normalized;
    }

    let remaining_width = (target_width - positive_total).max(0.0);
    let fallback_width = if remaining_width > 0.0 {
        remaining_width / non_positive_count as f64
    } else {
        positive_total / std::cmp::max(1, col_count - non_positive_count) as f64
    };

    normalized
        .into_iter()
        .map(|w| if w > 0.0 { w } else { fallback_width })
        .collect()
}

/// Resolve table column widths within a pixel budget.
pub fn resolve_table_column_widths(table_block: &TableBlock, content_width: f64) -> Vec<f64> {
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
        return resolve_autofit_column_widths(
            table_block,
            content_width,
            col_count,
            explicit_width_px,
        );
    }

    if !table_block.rows.is_empty() {
        column_widths = normalize_table_column_widths(&column_widths, col_count, target_width);
    }

    if !column_widths.is_empty()
        && let Some(explicit) = explicit_width_px
        && is_nonzero_number(explicit)
    {
        let total: f64 = column_widths.iter().fold(0.0, |sum, &w| sum + w);
        if total > 0.0 && (total - explicit).abs() > 1.0 {
            let scale = explicit / total;
            column_widths = column_widths.into_iter().map(|w| w * scale).collect();
        }
    }

    column_widths
}

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
    if is_nonzero_number(total) {
        return total;
    }
    if let Some(explicit) = explicit_width_px
        && is_nonzero_number(explicit)
    {
        return explicit;
    }
    content_width
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

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
    fn empty_array_returns_evenly_split_target_width() {
        assert_eq!(
            normalize_table_column_widths(&[], 3, 300.0),
            vec![100.0, 100.0, 100.0]
        );
    }

    #[test]
    fn missing_trailing_columns_inherit_average_of_existing_positives() {
        assert_eq!(
            normalize_table_column_widths(&[100.0, 200.0], 4, 1000.0),
            vec![100.0, 200.0, 150.0, 150.0]
        );
    }

    #[test]
    fn zero_negative_widths_split_the_leftover_target_evenly() {
        let out = normalize_table_column_widths(&[100.0, 0.0, 100.0, -5.0], 4, 400.0);
        assert_eq!(out[0], 100.0);
        assert_eq!(out[2], 100.0);
        assert_close_to(out[1], 100.0, 5);
        assert_close_to(out[3], 100.0, 5);
    }

    #[test]
    fn all_zero_returns_even_split_of_target() {
        assert_eq!(
            normalize_table_column_widths(&[0.0, 0.0, 0.0], 3, 300.0),
            vec![100.0, 100.0, 100.0]
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
}
