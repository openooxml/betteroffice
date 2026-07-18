//! Port of `packages/core/src/layout/pagination/tableRowBreak.ts`.
//!
//! Table row-break geometry. Word lets a table row break across a page
//! boundary; this module computes, per row, the set of safe break offsets
//! (the y of every line bottom across the row's content, including
//! vertically-merged cells that span into the row) so the paginator can snap
//! a break to the deepest whole line that still fits.
//!
//! Exported fns (1:1 with the TS module):
//! - `build_table_row_break_info` ← `buildTableRowBreakInfo(block, measure)`
//! - `snap_row_break` ← `snapRowBreak(info, rowIndex, fromOffset, maxSlice)`
//!
//! Consumes the spine's types (`types.rs`) directly. The place-loop
//! (`layout_table` in `hooks.rs`) uses this to fill `TableFragment`
//! rowStart/rowEnd + clipTop/clipBottom for mid-content row splits.

use serde::Serialize;

use crate::cell_layout::layout_cell_content;
use crate::table_grid::resolve_cell_grid;
use crate::types::{BlockExtent, LayoutBlock, TableBlock, TableExtent};

/// Per-table break geometry consumed by `snap_row_break`.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRowBreakInfo {
    /// Cumulative y of the top of each row; `row_tops[rows.len()]` is the table height.
    pub row_tops: Vec<f64>,
    /// Per-row sorted, de-duplicated line-bottom offsets (relative to the row
    /// top) at which a break is clean. Always includes the row's full height
    /// as the final boundary.
    pub break_offsets: Vec<Vec<f64>>,
}

/// SameValueZero membership insert (mirrors the JS `Set<number>`): NaN equals
/// NaN, +0 equals -0.
fn add_unique(offsets: &mut Vec<f64>, value: f64) {
    let same = |a: f64, b: f64| a == b || (a.is_nan() && b.is_nan());
    if !offsets.iter().any(|&existing| same(existing, value)) {
        offsets.push(value);
    }
}

fn cell_unbreakable_ranges(
    blocks: &[LayoutBlock],
    measures: &[BlockExtent],
    start_y: f64,
) -> Vec<(f64, f64)> {
    let mut ranges = Vec::new();
    let mut y = start_y;
    let mut previous_after = 0.0_f64;
    for (index, measure) in measures.iter().enumerate() {
        let block = blocks.get(index);
        if let (Some(LayoutBlock::Paragraph(paragraph)), BlockExtent::Paragraph(extent)) =
            (block, measure)
        {
            let spacing = paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.spacing.as_ref());
            y += previous_after.max(spacing.and_then(|value| value.before).unwrap_or(0.0));
            for line in &extent.lines {
                y += line.float_skip_before.unwrap_or(0.0);
                let top = y;
                y += line.line_height;
                ranges.push((top, y));
            }
            previous_after = spacing.and_then(|value| value.after).unwrap_or(0.0);
            continue;
        }
        let height = match measure {
            BlockExtent::Image(value) => Some(value.height),
            BlockExtent::TextBox(value) => Some(value.height),
            BlockExtent::Table(value) => Some(value.total_height),
            _ => None,
        };
        if let Some(height) = height {
            y += previous_after;
            let top = y;
            y += height;
            ranges.push((top, y));
            previous_after = 0.0;
        }
    }
    ranges
}

/// Resolves the cell grid once and collects, per row, every whole-line bottom
/// a break is allowed to snap to.
pub fn build_table_row_break_info(block: &TableBlock, measure: &TableExtent) -> TableRowBreakInfo {
    let row_count = measure.rows.len();
    // True (unrounded) cumulative row offsets — the paginator splits against
    // exact measured heights. The painter has a sibling `buildRowYPositions`
    // that rounds to whole pixels for crisp borders; keep the two SEPARATE
    // (don't "dedupe") or you break either break-offset alignment or crispness.
    let mut row_tops: Vec<f64> = Vec::with_capacity(row_count + 1);
    let mut acc = 0.0f64;
    for r in 0..row_count {
        row_tops.push(acc);
        acc += measure.rows[r].height;
    }
    row_tops.push(acc);

    // Use the shared grid resolution so "which cells cover row r" matches the
    // measurer and painter. A cell starting in row `sr` with rowSpan covers
    // rows [sr, sr + rowSpan); a merged cell spills its line bottoms into the
    // rows below its restart row.
    let resolved = resolve_cell_grid(block);
    let mut break_offsets: Vec<Vec<f64>> = Vec::with_capacity(row_count);
    for r in 0..row_count {
        let row_height = measure.rows[r].height;
        let mut offsets: Vec<f64> = Vec::new();
        let mut unbreakable_ranges: Vec<(f64, f64)> = Vec::new();
        add_unique(&mut offsets, row_height); // a row boundary is always a clean break

        for g in &resolved {
            // i64 arithmetic so a (theoretical) rowSpan of 0 mirrors the TS
            // `g.rowIndex + g.rowSpan - 1 < r` instead of underflowing.
            if g.row_index > r || (g.row_index + g.row_span) as i64 - 1 < r as i64 {
                continue;
            }
            let Some(source_cell) = block
                .rows
                .get(g.row_index)
                .and_then(|row| row.cells.get(g.cell_index))
            else {
                continue;
            };
            let Some(measured_cell) = measure
                .rows
                .get(g.row_index)
                .and_then(|row| row.cells.get(g.cell_index))
            else {
                continue;
            };
            // OOXML/TableNormal default top padding is 0 (matches measureTable).
            let pad_top = source_cell.padding.as_ref().map(|p| p.top).unwrap_or(0.0);
            let layout = layout_cell_content(
                Some(&source_cell.blocks),
                Some(&measured_cell.blocks),
                pad_top,
            );
            // Map cell-content y (relative to the cell/region top at
            // row_tops[start_row]) into this row's coordinate space
            // (relative to row_tops[r]).
            let shift = row_tops[r] - row_tops[g.row_index];
            for &b in &layout.flat_bottoms {
                let off = b - shift;
                if off > 0.0 && off < row_height {
                    add_unique(&mut offsets, off);
                }
            }
            for (top, bottom) in
                cell_unbreakable_ranges(&source_cell.blocks, &measured_cell.blocks, pad_top)
            {
                unbreakable_ranges.push((top - shift, bottom - shift));
            }
        }
        offsets.retain(|offset| {
            *offset == row_height
                || !unbreakable_ranges
                    .iter()
                    .any(|(top, bottom)| *offset > *top && *offset < *bottom)
        });
        offsets.sort_by(f64::total_cmp);
        break_offsets.push(offsets);
    }

    TableRowBreakInfo {
        row_tops,
        break_offsets,
    }
}

/// Given a row and how much of it has already been placed (`from_offset`),
/// return how many more px can be placed ending on a whole line, without
/// exceeding `max_slice`. Returns 0 when not even the first line fits.
pub fn snap_row_break(
    info: &TableRowBreakInfo,
    row_index: usize,
    from_offset: f64,
    max_slice: f64,
) -> f64 {
    let Some(offsets) = info.break_offsets.get(row_index) else {
        return 0.0;
    };
    if offsets.is_empty() {
        return 0.0;
    }
    let limit = from_offset + max_slice;
    let mut best = 0.0f64;
    for &off in offsets {
        if off <= from_offset {
            continue;
        }
        if off <= limit {
            best = off - from_offset;
        } else {
            break;
        }
    }
    best
}

// The TS module has no unit-test sibling — its only coverage is
// pagination/integration/table-row-break.test.ts, which drives the whole
// place loop (layoutDocument). These tests exercise the same 3-row
// merged-cell table geometry directly against buildTableRowBreakInfo /
// snapRowBreak; every expected value is verified against the TS
// implementation (bun run on identical input).
#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const LINE: f64 = 20.0;

    fn para() -> serde_json::Value {
        json!({ "kind": "paragraph", "id": 0, "runs": [] })
    }

    fn para_with_spacing(before: f64, after: f64) -> serde_json::Value {
        json!({
            "kind": "paragraph",
            "id": 0,
            "runs": [],
            "attrs": { "spacing": { "before": before, "after": after } },
        })
    }

    fn para_measure(lines: usize) -> serde_json::Value {
        let line = json!({
            "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
            "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": LINE,
        });
        json!({
            "kind": "paragraph",
            "lines": vec![line; lines],
            "totalHeight": lines as f64 * LINE,
        })
    }

    fn cell(row_span: Option<u32>, blocks: Vec<serde_json::Value>) -> serde_json::Value {
        json!({ "id": 0, "blocks": blocks, "rowSpan": row_span })
    }

    fn measured_cell(blocks: Vec<serde_json::Value>) -> serde_json::Value {
        json!({ "blocks": blocks, "width": 100.0, "height": 0.0 })
    }

    /// The integration-test table: 3 rows, 2 cols; col 0 is a rowSpan=3 merged
    /// cell with `merge_lines` of content; col 1 has one line per row. Row
    /// heights are supplied directly in the measure (as the real measurer
    /// would, after Word vmerge distribution).
    fn build_table(merge_lines: usize, row_heights: [f64; 3]) -> (TableBlock, TableExtent) {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [
                { "id": 0, "cells": [cell(Some(3), vec![para()]), cell(None, vec![para()])] },
                { "id": 1, "cells": [cell(None, vec![para()])] },
                { "id": 2, "cells": [cell(None, vec![para()])] },
            ],
            "columnWidths": [100.0, 100.0],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "columnWidths": [100.0, 100.0],
            "totalWidth": 200.0,
            "totalHeight": row_heights[0] + row_heights[1] + row_heights[2],
            "rows": [
                {
                    "height": row_heights[0],
                    "cells": [
                        measured_cell(vec![para_measure(merge_lines)]),
                        measured_cell(vec![para_measure(1)]),
                    ],
                },
                { "height": row_heights[1], "cells": [measured_cell(vec![para_measure(1)])] },
                { "height": row_heights[2], "cells": [measured_cell(vec![para_measure(1)])] },
            ],
        }))
        .unwrap();
        (block, measure)
    }

    #[test]
    fn builds_row_tops_and_whole_line_break_offsets_for_a_tall_merged_row() {
        // Last row holds the merged-cell overflow (Word distribution): 1, 1, 38 lines.
        let (block, measure) = build_table(40, [LINE, LINE, 38.0 * LINE]);
        let info = build_table_row_break_info(&block, &measure);

        assert_eq!(info.row_tops, vec![0.0, 20.0, 40.0, 800.0]);
        assert_eq!(info.break_offsets[0], vec![20.0]);
        assert_eq!(info.break_offsets[1], vec![20.0]);
        // The tall row: the merged cell's 37 in-row line bottoms plus the row
        // boundary — 20, 40, …, 760.
        assert_eq!(info.break_offsets[2].len(), 38);
        assert_eq!(info.break_offsets[2][0], 20.0);
        assert_eq!(*info.break_offsets[2].last().unwrap(), 760.0);
        // Break points are whole lines (multiples of the line height).
        for &off in &info.break_offsets[2] {
            assert_eq!(off % LINE, 0.0);
        }
    }

    #[test]
    fn snaps_a_break_to_the_deepest_whole_line_that_fits() {
        let (block, measure) = build_table(40, [LINE, LINE, 38.0 * LINE]);
        let info = build_table_row_break_info(&block, &measure);

        assert_eq!(snap_row_break(&info, 2, 0.0, 410.0), 400.0);
        assert_eq!(snap_row_break(&info, 2, 400.0, 410.0), 360.0);
        // Not even the first line fits.
        assert_eq!(snap_row_break(&info, 2, 0.0, 19.0), 0.0);
        assert_eq!(snap_row_break(&info, 0, 0.0, 100.0), 20.0);
    }

    #[test]
    fn shifts_merged_cell_bottoms_by_cell_padding_and_paragraph_spacing() {
        // rowSpan=2 merged cell with padTop 3 and spacing before 4 / after 6;
        // three 20px lines → bottoms at 27/47/67 in cell space.
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [
                {
                    "id": 0,
                    "cells": [{
                        "id": 0,
                        "rowSpan": 2,
                        "padding": { "top": 3.0, "right": 0.0, "bottom": 0.0, "left": 0.0 },
                        "blocks": [para_with_spacing(4.0, 6.0)],
                    }],
                },
                { "id": 1, "cells": [] },
            ],
            "columnWidths": [100.0],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "columnWidths": [100.0],
            "totalWidth": 100.0,
            "totalHeight": 50.0,
            "rows": [
                { "height": 30.0, "cells": [measured_cell(vec![para_measure(3)])] },
                { "height": 20.0, "cells": [] },
            ],
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.row_tops, vec![0.0, 30.0, 50.0]);
        // Row 0 keeps the in-row bottom 27; row 1 sees 47-30=17 from the spill.
        assert_eq!(info.break_offsets, vec![vec![27.0, 30.0], vec![17.0, 20.0]]);
    }

    #[test]
    fn snap_returns_zero_for_a_row_without_offsets() {
        let info = TableRowBreakInfo {
            row_tops: vec![0.0],
            break_offsets: vec![],
        };
        assert_eq!(snap_row_break(&info, 0, 0.0, 100.0), 0.0);
        assert_eq!(snap_row_break(&info, 5, 0.0, 100.0), 0.0);
    }

    #[test]
    fn rejects_a_boundary_that_would_slice_a_line_in_another_cell() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "blocks": [para()] },
                { "id": 1, "blocks": [para()] }
            ] }],
            "columnWidths": [100, 100],
        }))
        .unwrap();
        let line20 = para_measure(2);
        let line30 = json!({
            "kind": "paragraph",
            "lines": [
                { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                  "width": 0, "ascent": 0, "descent": 0, "lineHeight": 30 },
                { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                  "width": 0, "ascent": 0, "descent": 0, "lineHeight": 30 }
            ],
            "totalHeight": 60,
        });
        let measure: TableExtent = serde_json::from_value(json!({
            "rows": [{ "height": 60, "cells": [
                measured_cell(vec![line20]), measured_cell(vec![line30])
            ] }],
            "columnWidths": [100, 100], "totalWidth": 200, "totalHeight": 60,
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![60.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 40.0), 0.0);
    }
}
