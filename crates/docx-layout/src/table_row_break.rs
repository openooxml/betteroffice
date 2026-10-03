//! Whole-line table row-break geometry.

use std::cell::OnceCell;

use serde::Serialize;

use crate::cell_layout::{
    cell_vertical_offset, is_floating_image, layout_cell_content, nested_table_float_offset,
};
use crate::keep_together::{paragraph_is_unbreakable, paragraph_widow_control};
use crate::table_grid::resolve_cell_grid;
use crate::types::{
    BlockExtent, CellClip, LayoutBlock, Run, TableBlock, TableCell, TableCellExtent, TableExtent,
};

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

/// Inserts with SameValueZero semantics: NaN equals NaN and +0 equals -0.
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
    paragraph_rules: bool,
) -> Vec<(f64, f64)> {
    let mut ranges = Vec::new();
    let mut y = start_y;
    let mut previous_after = 0.0_f64;
    let mut keep_next_top: Option<f64> = None;
    for (index, measure) in measures.iter().enumerate() {
        let preceding_keep_top = keep_next_top.take();
        let block = blocks.get(index);
        if let (Some(LayoutBlock::Paragraph(paragraph)), BlockExtent::Paragraph(extent)) =
            (block, measure)
        {
            let spacing = paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.spacing.as_ref());
            y += previous_after.max(spacing.and_then(|value| value.before).unwrap_or(0.0));
            let first = ranges.len();
            for line in &extent.lines {
                y += line.float_skip_before.unwrap_or(0.0);
                let top = y;
                y += line.line_height;
                ranges.push((top, y));
            }
            let lines = &ranges[first..];
            let first_bottom = lines.first().map(|&(_, bottom)| bottom);
            let last_top = lines.last().map(|&(top, _)| top);
            if paragraph_rules
                && let (Some(&(top, _)), Some(&(_, bottom))) = (lines.first(), lines.last())
            {
                if paragraph_is_unbreakable(paragraph, extent) {
                    ranges.push((top, bottom));
                } else if paragraph_widow_control(paragraph, extent) {
                    let (second_bottom, penultimate_top) = (lines[1].1, lines[lines.len() - 2].0);
                    ranges.push((top, second_bottom));
                    ranges.push((penultimate_top, bottom));
                }
            }
            if paragraph_rules {
                if let (Some(top), Some(bottom)) = (preceding_keep_top, first_bottom) {
                    ranges.push((top, bottom));
                }
                keep_next_top =
                    if paragraph.attrs.as_ref().and_then(|attrs| attrs.keep_next) == Some(true) {
                        last_top
                    } else {
                        None
                    };
            }
            previous_after = spacing.and_then(|value| value.after).unwrap_or(0.0);
            continue;
        }
        let height = match measure {
            BlockExtent::Image(value) => Some(value.height),
            BlockExtent::TextBox(value) => Some(value.height),
            BlockExtent::Table(value) => Some(value.total_height),
            BlockExtent::Shape(value) => Some(value.height),
            BlockExtent::Chart(value) => Some(value.height),
            _ => None,
        };
        if let Some(LayoutBlock::Shape(shape)) = block
            && crate::cell_layout::cell_overlay_drawing(
                shape.position.is_some(),
                shape.wrap_type.as_deref(),
            )
        {
            continue;
        }
        if let Some(height) = height {
            y += previous_after;
            if let Some(LayoutBlock::Table(table)) = block
                && let Some(offset) = nested_table_float_offset(table.floating.as_ref())
            {
                ranges.push((y + offset, y + offset + height));
                previous_after = 0.0;
                continue;
            }
            let top = y;
            y += height;
            ranges.push((top, y));
            previous_after = 0.0;
        }
    }
    ranges
}

/// Resolves the cell grid once and collects, per row, every whole-line bottom
/// a break is allowed to snap to, where widow/orphan control and `keepLines`
/// allow one.
pub fn build_table_row_break_info(block: &TableBlock, measure: &TableExtent) -> TableRowBreakInfo {
    row_break_info(block, measure, true)
}

/// A table's row break geometry under the paragraph rules, with every
/// whole-line bottom built on first use for a row those rules leave no break
/// in a whole column, which Word then breaks at any line. A floating table
/// keeps whole-line breaks throughout.
/// Whether row `r`'s `w:trHeight` minimum, not its cells' content, sets its
/// measured height. A row a vertically merged cell covers can grow with that
/// cell's content, so it keeps its line breaks.
fn minimum_height_governs(
    block: &TableBlock,
    measure: &TableExtent,
    resolved: &[crate::table_grid::ResolvedGridCell],
    r: usize,
) -> bool {
    let Some(minimum) = block
        .rows
        .get(r)
        .and_then(|row| row.height)
        .filter(|height| *height > 0.0)
    else {
        return false;
    };
    if resolved
        .iter()
        .any(|grid| grid.row_span > 1 && grid.row_index <= r && r < grid.row_index + grid.row_span)
    {
        return false;
    }
    let (mut content, mut padding) = (0.0_f64, 0.0_f64);
    for grid in resolved.iter().filter(|grid| grid.row_index == r) {
        let (Some(cell), Some(measured)) = (
            block.rows[r].cells.get(grid.cell_index),
            measure.rows[r].cells.get(grid.cell_index),
        ) else {
            continue;
        };
        content = content.max(measured.height);
        padding = padding.max(
            cell.padding
                .as_ref()
                .map_or(0.0, |padding| padding.top + padding.bottom),
        );
    }
    content > 0.0 && minimum + padding >= content
}

fn cell_has_drawing(blocks: &[LayoutBlock]) -> bool {
    blocks.iter().any(|block| match block {
        LayoutBlock::Paragraph(paragraph) => paragraph.runs.iter().any(|run| {
            matches!(run, Run::Image(image) if image.position.is_some()
                || is_floating_image(image.wrap_type.as_deref(), image.display_mode.as_deref()))
        }),
        LayoutBlock::Table(table) => {
            table.floating.is_some()
                || table
                    .rows
                    .iter()
                    .any(|row| row.cells.iter().any(|cell| cell_has_drawing(&cell.blocks)))
        }
        _ => true,
    })
}

pub(crate) struct RowBreaks<'a> {
    block: &'a TableBlock,
    measure: &'a TableExtent,
    pub(crate) kept: TableRowBreakInfo,
    lines: OnceCell<TableRowBreakInfo>,
    cells: OnceCell<Vec<Option<Vec<CellBreaks>>>>,
}

impl<'a> RowBreaks<'a> {
    pub(crate) fn new(block: &'a TableBlock, measure: &'a TableExtent) -> Self {
        Self {
            block,
            measure,
            kept: row_break_info(block, measure, block.floating.is_none()),
            lines: OnceCell::new(),
            cells: OnceCell::new(),
        }
    }

    pub(crate) fn lines(&self) -> &TableRowBreakInfo {
        self.lines
            .get_or_init(|| row_break_info(self.block, self.measure, false))
    }

    fn cells(&self, row: usize) -> Option<&[CellBreaks]> {
        self.cells
            .get_or_init(|| {
                let resolved = resolve_cell_grid(self.block);
                self.block
                    .rows
                    .iter()
                    .enumerate()
                    .map(|(r, row)| {
                        let measured = self.measure.rows.get(r)?;
                        if row.cant_split == Some(true)
                            || row.is_exact_height()
                            || row.is_header == Some(true)
                            || minimum_height_governs(self.block, self.measure, &resolved, r)
                            || resolved.iter().any(|g| {
                                g.row_span > 1 && g.row_index <= r && r < g.row_index + g.row_span
                            })
                            || row.cells.iter().any(|cell| {
                                cell.row_span.is_some_and(|span| span != 1.0)
                                    || !matches!(cell.vertical_align.as_deref(), None | Some("top"))
                                    || !matches!(
                                        cell.text_direction.as_deref(),
                                        None | Some("lrTb")
                                    )
                                    || !crate::footnotes::collect_note_refs(&cell.blocks).is_empty()
                                    || cell_has_drawing(&cell.blocks)
                            })
                        {
                            return None;
                        }
                        let mut cells: Vec<_> = row
                            .cells
                            .iter()
                            .enumerate()
                            .map(|(index, cell)| {
                                let measure = measured.cells.get(index)?;
                                let kept = cell_row_geometry(
                                    cell,
                                    measure,
                                    r,
                                    1,
                                    r,
                                    &self.kept.row_tops,
                                    self.block.floating.is_none(),
                                );
                                let lines = cell_row_geometry(
                                    cell,
                                    measure,
                                    r,
                                    1,
                                    r,
                                    &self.kept.row_tops,
                                    false,
                                );
                                let end = kept.end.min(measured.height).max(0.0);
                                Some(CellBreaks {
                                    kept: cell_offsets(&kept, end),
                                    lines: cell_offsets(&lines, end),
                                    end,
                                })
                            })
                            .collect::<Option<Vec<_>>>()?;
                        if let Some(cell) = cells.iter_mut().max_by(|a, b| a.end.total_cmp(&b.end))
                            && cell.end < measured.height
                        {
                            cell.kept.retain(|offset| *offset != cell.end);
                            cell.lines.retain(|offset| *offset != cell.end);
                            cell.end = measured.height;
                            cell.kept.push(cell.end);
                            cell.lines.push(cell.end);
                        }
                        Some(cells)
                    })
                    .collect()
            })
            .get(row)?
            .as_deref()
    }

    pub(crate) fn cell_remaining(&self, row: usize, tops: &[f64]) -> f64 {
        self.cells(row)
            .into_iter()
            .flatten()
            .zip(tops)
            .map(|(cell, top)| (cell.end - top).max(0.0))
            .fold(0.0, f64::max)
    }

    pub(crate) fn first_cell_slice(&self, row: usize, consumed: f64, capacity: f64) -> Option<f64> {
        let minimum = self
            .cells(row)?
            .iter()
            .filter_map(|cell| {
                let top = consumed.min(cell.end);
                cell.kept
                    .iter()
                    .copied()
                    .find(|offset| *offset > top)
                    .map(|offset| offset - top)
            })
            .min_by(f64::total_cmp)?;
        let shared = snap_row_break(&self.kept, row, consumed, minimum);
        (minimum <= capacity
            && self
                .cell_slice(row, consumed, None, minimum, shared, false, capacity)
                .is_some())
        .then_some(minimum)
    }

    pub(crate) fn cell_slice(
        &self,
        row: usize,
        consumed: f64,
        tops: Option<&[f64]>,
        budget: f64,
        shared_slice: f64,
        whole_lines: bool,
        capacity: f64,
    ) -> Option<CellRowSlice> {
        let cells = self.cells(row)?;
        let last_fitting = |offsets: &[f64], top: f64| {
            offsets
                .iter()
                .copied()
                .rfind(|offset| *offset > top && *offset <= top + budget)
                .unwrap_or(top)
        };
        let mut clips: Vec<_> = cells
            .iter()
            .enumerate()
            .map(|(index, cell)| {
                let top = tops
                    .and_then(|tops| tops.get(index))
                    .copied()
                    .unwrap_or(consumed)
                    .min(cell.end);
                let offsets = if whole_lines { &cell.lines } else { &cell.kept };
                CellClip {
                    row,
                    cell: index,
                    top,
                    bottom: last_fitting(offsets, top),
                }
            })
            .collect();
        let height = |clips: &[CellClip]| {
            clips
                .iter()
                .map(|clip| clip.bottom - clip.top)
                .fold(0.0, f64::max)
        };
        if height(&clips) <= 0.0
            || (tops.is_none()
                && !cells.iter().zip(&clips).any(|(cell, clip)| {
                    cell.end > consumed + shared_slice && clip.bottom > consumed + shared_slice
                }))
        {
            return None;
        }
        // A cell its paragraph rules leave no break even on a fresh page cuts at
        // whole lines, as an oversized row does, while the other cells progress.
        for (cell, clip) in cells.iter().zip(&mut clips) {
            if !whole_lines && clip.bottom == clip.top && cell.kept_oversized(clip.top, capacity) {
                clip.bottom = last_fitting(&cell.lines, clip.top);
            }
        }
        let height = height(&clips);
        let complete = cells
            .iter()
            .zip(&clips)
            .all(|(cell, clip)| clip.bottom == cell.end);
        Some(CellRowSlice {
            clips,
            height,
            complete,
        })
    }

    pub(crate) fn cell_fresh_slice(&self, row: usize, tops: &[f64], capacity: f64) -> f64 {
        let minimum = |whole_lines| {
            self.cells(row)
                .into_iter()
                .flatten()
                .zip(tops)
                .filter_map(|(cell, top)| {
                    let offsets = if whole_lines { &cell.lines } else { &cell.kept };
                    offsets
                        .iter()
                        .copied()
                        .find(|offset| *offset > *top)
                        .map(|offset| offset - top)
                })
                .min_by(f64::total_cmp)
                .unwrap_or(0.0)
        };
        let kept = minimum(false);
        if kept > capacity { minimum(true) } else { kept }
    }

    pub(crate) fn cell_remainder(&self, row: usize, tops: &[f64]) -> CellRowSlice {
        let clips: Vec<_> = self
            .cells(row)
            .into_iter()
            .flatten()
            .zip(tops)
            .enumerate()
            .map(|(cell, (info, top))| CellClip {
                row,
                cell,
                top: *top,
                bottom: info.end,
            })
            .collect();
        CellRowSlice {
            clips,
            height: self.cell_remaining(row, tops),
            complete: true,
        }
    }

    /// Whether the paragraph rules alone leave `row` no break from `consumed`
    /// on in a column `capacity` tall.
    pub(crate) fn kept_oversized(&self, row: usize, consumed: f64, capacity: f64) -> bool {
        let kept = minimum_break_slice(self.measure, &self.kept, row, consumed);
        kept > capacity && minimum_break_slice(self.measure, self.lines(), row, consumed) < kept
    }

    /// The smallest slice of `row` from `consumed` on that a fresh column
    /// `capacity` tall places.
    pub(crate) fn fresh_slice(&self, row: usize, consumed: f64, capacity: f64) -> f64 {
        if self
            .block
            .rows
            .get(row)
            .is_some_and(|row| row.cant_split.unwrap_or(false))
        {
            return minimum_row_slice(self.block, self.measure, self.lines(), row, consumed);
        }
        let info = if self.kept_oversized(row, consumed, capacity) {
            self.lines()
        } else {
            &self.kept
        };
        minimum_break_slice(self.measure, info, row, consumed)
    }
}

struct CellBreaks {
    kept: Vec<f64>,
    lines: Vec<f64>,
    end: f64,
}

impl CellBreaks {
    /// Whether the paragraph rules leave this cell no break from `top` on in a
    /// fresh column `capacity` tall while whole lines still break it sooner, so
    /// the cell falls back to whole lines as an oversized row does.
    fn kept_oversized(&self, top: f64, capacity: f64) -> bool {
        let next = |offsets: &[f64]| {
            offsets
                .iter()
                .copied()
                .find(|offset| *offset > top)
                .map(|offset| offset - top)
        };
        matches!((next(&self.kept), next(&self.lines)), (Some(kept), Some(line)) if kept > capacity && line < kept)
    }
}

pub(crate) struct CellRowSlice {
    pub(crate) clips: Vec<CellClip>,
    pub(crate) height: f64,
    pub(crate) complete: bool,
}

fn cell_offsets(geometry: &CellRowGeometry, end: f64) -> Vec<f64> {
    let mut offsets: Vec<_> = geometry
        .offsets
        .iter()
        .copied()
        .filter(|offset| {
            *offset > 0.0
                && *offset < end
                && geometry.ranges.iter().any(|(_, bottom)| *bottom > *offset)
                && !geometry
                    .ranges
                    .iter()
                    .any(|(top, bottom)| *offset > *top && *offset < *bottom)
        })
        .collect();
    add_unique(&mut offsets, end);
    offsets.sort_by(f64::total_cmp);
    offsets
}

struct CellRowGeometry {
    offsets: Vec<f64>,
    ranges: Vec<(f64, f64)>,
    end: f64,
}

fn cell_row_geometry(
    cell: &TableCell,
    measured: &TableCellExtent,
    source_row: usize,
    row_span: usize,
    r: usize,
    row_tops: &[f64],
    paragraph_rules: bool,
) -> CellRowGeometry {
    // OOXML TableNormal defaults top padding to zero.
    let pad_top = cell.padding.as_ref().map(|p| p.top).unwrap_or(0.0);
    let pad_bottom = cell.padding.as_ref().map(|p| p.bottom).unwrap_or(0.0);
    let border_width = |edge: Option<&crate::types::CellBorderSpec>| {
        edge.filter(|edge| !matches!(edge.style.as_deref(), Some("none" | "nil")))
            .map_or(0.0, |edge| edge.width.unwrap_or(1.0))
    };
    let border_top = if source_row == 0 {
        border_width(cell.borders.as_ref().and_then(|b| b.top.as_ref()))
    } else {
        0.0
    };
    let border_bottom = border_width(cell.borders.as_ref().and_then(|b| b.bottom.as_ref()));
    let layout = layout_cell_content(Some(&cell.blocks), Some(&measured.blocks), pad_top);
    let cell_end = (source_row + row_span).min(row_tops.len() - 1);
    let cell_height = row_tops[cell_end] - row_tops[source_row];
    let content_offset = border_top
        + cell_vertical_offset(
            cell.vertical_align.as_deref(),
            cell_height,
            measured.height,
            layout.content_height,
            pad_top + border_top,
            pad_bottom + border_bottom,
        );
    let shift = row_tops[r] - row_tops[source_row];

    CellRowGeometry {
        offsets: layout
            .flat_bottoms
            .iter()
            .map(|b| b + content_offset - shift)
            .collect(),
        ranges: cell_unbreakable_ranges(&cell.blocks, &measured.blocks, pad_top, paragraph_rules)
            .into_iter()
            .map(|(top, bottom)| {
                (
                    top + content_offset - shift,
                    bottom + content_offset - shift,
                )
            })
            .collect(),
        end: pad_top + content_offset + layout.content_height + pad_bottom + border_bottom - shift,
    }
}

fn row_break_info(
    block: &TableBlock,
    measure: &TableExtent,
    paragraph_rules: bool,
) -> TableRowBreakInfo {
    let row_count = measure.rows.len();
    // Pagination uses unrounded row offsets; border painting rounds separately.
    let mut row_tops: Vec<f64> = Vec::with_capacity(row_count + 1);
    let mut acc = 0.0f64;
    for r in 0..row_count {
        row_tops.push(acc);
        acc += measure.rows[r].height;
    }
    row_tops.push(acc);

    // Use the shared grid resolution so "which cells cover row r" agrees with
    // measurement and paint. A cell starting in row `sr` with rowSpan covers
    // rows [sr, sr + rowSpan); a merged cell spills its line bottoms into the
    // rows below its restart row.
    let resolved = resolve_cell_grid(block);
    let mut break_offsets: Vec<Vec<f64>> = Vec::with_capacity(row_count);
    for r in 0..row_count {
        let row_height = measure.rows[r].height;
        // Word treats an exact-height row as an indivisible fixed block
        // (measurement takes its height verbatim), so it offers only its
        // full-height boundary. Filtering here — rather than adding a second
        // check in the paginator — makes every downstream consumer
        // (`snap_row_break`, `minimum_row_slice`, `first_table_fragment_height`)
        // see the row as atomic by construction.
        // A row its minimum height sizes, taller than its content, moves whole
        // too: Word breaks a row only inside content that overflows it.
        if block.rows.get(r).is_some_and(|row| row.is_exact_height())
            || minimum_height_governs(block, measure, &resolved, r)
        {
            break_offsets.push(vec![row_height]);
            continue;
        }
        let mut offsets: Vec<f64> = Vec::new();
        let mut unbreakable_ranges: Vec<(f64, f64)> = Vec::new();
        add_unique(&mut offsets, row_height); // a row boundary is always a clean break

        for g in &resolved {
            // Signed arithmetic avoids underflow for a theoretical zero row span.
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
            let geometry = cell_row_geometry(
                source_cell,
                measured_cell,
                g.row_index,
                g.row_span,
                r,
                &row_tops,
                paragraph_rules,
            );
            for off in geometry.offsets {
                if off > 0.0 && off < row_height {
                    add_unique(&mut offsets, off);
                }
            }
            unbreakable_ranges.extend(geometry.ranges);
        }
        offsets.retain(|offset| {
            *offset == row_height
                || (unbreakable_ranges
                    .iter()
                    .any(|(_, bottom)| *bottom > *offset)
                    && !unbreakable_ranges
                        .iter()
                        .any(|(top, bottom)| *offset > *top && *offset < *bottom))
        });
        offsets.sort_by(f64::total_cmp);
        break_offsets.push(offsets);
    }

    TableRowBreakInfo {
        row_tops,
        break_offsets,
    }
}

pub(crate) fn minimum_row_slice(
    block: &TableBlock,
    measure: &TableExtent,
    info: &TableRowBreakInfo,
    row: usize,
    consumed: f64,
) -> f64 {
    let remaining = measure.rows[row].height - consumed;
    if consumed == 0.0
        && block
            .rows
            .get(row)
            .is_some_and(|row| row.cant_split.unwrap_or(false))
    {
        return remaining;
    }
    minimum_break_slice(measure, info, row, consumed)
}

fn minimum_break_slice(
    measure: &TableExtent,
    info: &TableRowBreakInfo,
    row: usize,
    consumed: f64,
) -> f64 {
    let remaining = measure.rows[row].height - consumed;
    info.break_offsets[row]
        .iter()
        .copied()
        .find(|offset| *offset > consumed)
        .map_or(remaining, |offset| offset - consumed)
}

pub(crate) fn first_table_fragment_height(
    block: &TableBlock,
    measure: &TableExtent,
    info: &TableRowBreakInfo,
) -> f64 {
    let headers = block
        .rows
        .iter()
        .take_while(|row| row.is_header.unwrap_or(false))
        .count()
        .min(measure.rows.len());
    if headers == 0 {
        return measure.rows.first().map_or(0.0, |row| row.height);
    }
    let header_height: f64 = measure.rows[..headers].iter().map(|row| row.height).sum();
    header_height
        + if headers < measure.rows.len() {
            minimum_row_slice(block, measure, info, headers, 0.0)
        } else {
            0.0
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

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const LINE: f64 = 20.0;

    fn para() -> serde_json::Value {
        json!({ "kind": "paragraph", "id": 0, "runs": [], "attrs": { "widowControl": false } })
    }

    fn para_with_spacing(before: f64, after: f64) -> serde_json::Value {
        json!({
            "kind": "paragraph",
            "id": 0,
            "runs": [],
            "attrs": {
                "spacing": { "before": before, "after": after },
                "widowControl": false,
            },
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

    fn single_cell_table(
        blocks: Vec<serde_json::Value>,
        measures: Vec<serde_json::Value>,
        row_height: f64,
    ) -> (TableBlock, TableExtent) {
        let block = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [cell(None, blocks)] }],
            "columnWidths": [100.0],
        }))
        .unwrap();
        let measure = serde_json::from_value(json!({
            "columnWidths": [100.0],
            "totalWidth": 100.0,
            "totalHeight": row_height,
            "rows": [{ "height": row_height, "cells": [measured_cell(measures)] }],
        }))
        .unwrap();
        (block, measure)
    }

    #[test]
    fn keep_next_removes_only_the_paragraph_boundary() {
        for keep_next in [Some(true), Some(false), None] {
            let mut middle = para();
            if let Some(keep_next) = keep_next {
                middle["attrs"]["keepNext"] = json!(keep_next);
            }
            let (block, measure) =
                single_cell_table(vec![para(), middle, para()], vec![para_measure(1); 3], 60.0);
            let info = build_table_row_break_info(&block, &measure);
            let (offsets, snapped) = if keep_next == Some(true) {
                (vec![20.0, 60.0], 20.0)
            } else {
                (vec![20.0, 40.0, 60.0], 40.0)
            };
            assert_eq!(info.break_offsets[0], offsets);
            assert_eq!(snap_row_break(&info, 0, 0.0, 40.0), snapped);
        }
    }

    #[test]
    fn keep_next_allows_splitting_both_paragraphs() {
        let mut kept = para();
        kept["attrs"]["keepNext"] = json!(true);
        let (block, measure) =
            single_cell_table(vec![para(), kept, para()], vec![para_measure(2); 3], 120.0);
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![20.0, 40.0, 60.0, 100.0, 120.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 60.0), 60.0);
        assert_eq!(snap_row_break(&info, 0, 0.0, 100.0), 100.0);
    }

    #[test]
    fn keep_next_composes_with_follower_widow_control() {
        for widow_control in [Some(true), None] {
            for (lines, offsets) in [(2, vec![60.0]), (3, vec![80.0]), (4, vec![60.0, 100.0])] {
                let mut kept = para();
                kept["attrs"]["keepNext"] = json!(true);
                let mut follower = para();
                follower["attrs"]
                    .as_object_mut()
                    .unwrap()
                    .remove("widowControl");
                if let Some(widow_control) = widow_control {
                    follower["attrs"]["widowControl"] = json!(widow_control);
                }
                let (block, measure) = single_cell_table(
                    vec![kept, follower],
                    vec![para_measure(1), para_measure(lines)],
                    (lines + 1) as f64 * LINE,
                );
                let info = build_table_row_break_info(&block, &measure);
                assert_eq!(info.break_offsets[0], offsets);
                assert_eq!(snap_row_break(&info, 0, 0.0, 40.0), 0.0);
                if lines == 4 {
                    assert_eq!(snap_row_break(&info, 0, 0.0, 60.0), 60.0);
                }
            }
        }
    }

    #[test]
    fn keep_next_stops_at_nested_tables() {
        let mut kept = para();
        kept["attrs"]["keepNext"] = json!(true);
        let (nested, nested_measure) = single_cell_table(vec![para()], vec![para_measure(2)], 40.0);
        let mut nested = serde_json::to_value(nested).unwrap();
        nested["kind"] = json!("table");
        let mut nested_measure = serde_json::to_value(nested_measure).unwrap();
        nested_measure["kind"] = json!("table");
        let (block, measure) = single_cell_table(
            vec![kept, nested, para()],
            vec![para_measure(1), nested_measure, para_measure(1)],
            80.0,
        );
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![20.0, 60.0, 80.0]);
    }

    #[test]
    fn floating_table_ignores_local_keep_next() {
        let mut kept = para();
        kept["attrs"]["keepNext"] = json!(true);
        let (mut block, measure) =
            single_cell_table(vec![para(), kept, para()], vec![para_measure(1); 3], 60.0);
        block.floating = Some(serde_json::from_value(json!({})).unwrap());
        let breaks = RowBreaks::new(&block, &measure);
        assert_eq!(&breaks.kept, breaks.lines());
        assert_eq!(breaks.kept.break_offsets[0], vec![20.0, 40.0, 60.0]);
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
    fn keeps_the_last_line_with_trailing_cell_padding_and_paragraph_spacing() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [{
                "id": 0,
                "padding": { "top": 5, "bottom": 5, "left": 0, "right": 0 },
                "blocks": [para_with_spacing(0.0, 3.0)]
            }] }],
            "columnWidths": [100],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "rows": [{ "height": 53, "cells": [{
                "blocks": [para_measure(2)], "width": 100, "height": 53
            }] }],
            "columnWidths": [100], "totalWidth": 100, "totalHeight": 53,
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(snap_row_break(&info, 0, 0.0, 50.0), 25.0);
        assert_eq!(snap_row_break(&info, 0, 25.0, 27.0), 0.0);
        assert_eq!(snap_row_break(&info, 0, 25.0, 28.0), 28.0);
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
        let breaks = RowBreaks::new(&block, &measure);
        let slice = breaks
            .cell_slice(0, 0.0, None, 40.0, 0.0, false, f64::INFINITY)
            .unwrap();
        assert_eq!(slice.height, 40.0);
        assert_eq!(
            slice.clips,
            vec![
                CellClip {
                    row: 0,
                    cell: 0,
                    top: 0.0,
                    bottom: 40.0
                },
                CellClip {
                    row: 0,
                    cell: 1,
                    top: 0.0,
                    bottom: 30.0
                },
            ]
        );
        let tops: Vec<_> = slice.clips.iter().map(|clip| clip.bottom).collect();
        let continuation = breaks.cell_remainder(0, &tops);
        assert_eq!(continuation.height, 30.0);
        assert_eq!(continuation.clips[0].top, continuation.clips[0].bottom);
        assert_eq!(continuation.clips[1].top, 30.0);
        assert_eq!(continuation.clips[1].bottom, 60.0);
    }

    #[test]
    fn keeps_centered_cell_text_whole_across_a_page_break() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "verticalAlign": "center", "blocks": [para()] },
                { "id": 1, "blocks": [para()] }
            ] }],
            "columnWidths": [100, 100],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "rows": [{ "height": 40, "cells": [
                { "blocks": [para_measure(1)], "width": 100, "height": 20 },
                { "blocks": [para_measure(2)], "width": 100, "height": 40 }
            ] }],
            "columnWidths": [100, 100], "totalWidth": 200, "totalHeight": 40,
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![40.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 30.0), 0.0);
    }

    #[test]
    fn offsets_bottom_aligned_text_inside_a_tall_row() {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [
                { "id": 0, "verticalAlign": "bottom", "blocks": [para()] }
            ] }],
            "columnWidths": [100],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "rows": [{ "height": 60, "cells": [
                { "blocks": [para_measure(1)], "width": 100, "height": 20 }
            ] }],
            "columnWidths": [100], "totalWidth": 100, "totalHeight": 60,
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![60.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 40.0), 0.0);
    }

    #[test]
    fn preserves_aligned_line_boundaries_with_fractional_padding() {
        let padding = 1.0 / 15.0;
        let line_height = 17.89453125;
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{ "id": 0, "cells": [{
                "id": 0,
                "verticalAlign": "center",
                "padding": { "top": padding, "bottom": padding, "left": 0, "right": 0 },
                "blocks": [para()]
            }] }],
            "columnWidths": [100],
        }))
        .unwrap();
        let mut paragraph = para_measure(2);
        for line in paragraph["lines"].as_array_mut().unwrap() {
            line["lineHeight"] = json!(line_height);
        }
        paragraph["totalHeight"] = json!(2.0 * line_height);
        let measure: TableExtent = serde_json::from_value(json!({
            "rows": [{ "height": 100, "cells": [{
                "blocks": [paragraph], "width": 100, "height": 2.0 * (line_height + padding)
            }] }],
            "columnWidths": [100], "totalWidth": 100, "totalHeight": 100,
        }))
        .unwrap();
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(snap_row_break(&info, 0, 0.0, 50.0), 50.0);
    }

    fn single_row_table(
        height: Option<f64>,
        height_rule: Option<&str>,
        cant_split: Option<bool>,
        lines: usize,
        row_height: f64,
    ) -> (TableBlock, TableExtent) {
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [{
                "id": 0,
                "height": height,
                "heightRule": height_rule,
                "cantSplit": cant_split,
                "cells": [{ "id": 0, "blocks": [para()] }],
            }],
            "columnWidths": [100.0],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "columnWidths": [100.0],
            "totalWidth": 100.0,
            "totalHeight": row_height,
            "rows": [{
                "height": row_height,
                "cells": [measured_cell(vec![para_measure(lines)])],
            }],
        }))
        .unwrap();
        (block, measure)
    }

    #[test]
    fn a_row_its_minimum_height_sizes_offers_only_its_full_height_boundary() {
        for (lines, atomic) in [(2, true), (3, true), (4, false)] {
            let content = lines as f64 * LINE;
            let (block, mut measure) =
                single_row_table(Some(3.0 * LINE), None, None, lines, content.max(3.0 * LINE));
            measure.rows[0].cells[0].height = content;
            let info = build_table_row_break_info(&block, &measure);
            assert_eq!(info.break_offsets[0].len() == 1, atomic, "{lines} lines");
        }
    }

    #[test]
    fn a_minimum_height_row_a_merged_cell_covers_keeps_its_line_breaks() {
        let (mut block, mut measure) = build_table(40, [20.0, 20.0, 760.0]);
        block.rows[2].height = Some(60.0);
        measure.rows[2].cells[0].height = LINE;
        let info = build_table_row_break_info(&block, &measure);
        assert!(info.break_offsets[2].len() > 1);
    }

    #[test]
    fn exact_row_offers_only_its_full_height_boundary() {
        let (block, measure) = single_row_table(Some(60.0), Some("exact"), None, 3, 60.0);
        assert!(block.rows[0].is_exact_height());
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![60.0]);
        // A partial budget fits no clean boundary; the full budget fits whole.
        assert_eq!(snap_row_break(&info, 0, 0.0, 30.0), 0.0);
        assert_eq!(snap_row_break(&info, 0, 0.0, 60.0), 60.0);
        assert_eq!(minimum_row_slice(&block, &measure, &info, 0, 0.0), 60.0);
    }

    #[test]
    fn at_least_row_keeps_interior_line_boundaries() {
        let (block, measure) = single_row_table(Some(20.0), Some("atLeast"), None, 3, 60.0);
        assert!(!block.rows[0].is_exact_height());
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![20.0, 40.0, 60.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 30.0), 20.0);
    }

    #[test]
    fn absent_rule_keeps_interior_line_boundaries() {
        let (block, measure) = single_row_table(None, None, None, 3, 60.0);
        assert!(!block.rows[0].is_exact_height());
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![20.0, 40.0, 60.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 30.0), 20.0);
    }

    #[test]
    fn exact_and_cant_split_together_stay_atomic() {
        let (block, measure) = single_row_table(Some(60.0), Some("exact"), Some(true), 3, 60.0);
        assert!(block.rows[0].is_exact_height());
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![60.0]);
        assert_eq!(snap_row_break(&info, 0, 0.0, 30.0), 0.0);
        assert_eq!(snap_row_break(&info, 0, 0.0, 60.0), 60.0);
        assert_eq!(minimum_row_slice(&block, &measure, &info, 0, 0.0), 60.0);
    }

    #[test]
    fn exact_row_spanned_by_a_merged_cell_stays_atomic() {
        // Row 1 is exact but covered by a rowSpan=2 merged cell starting in row
        // 0. The exact row must ignore the spill interior and offer only its
        // full-height boundary, while the non-exact row above is unaffected.
        let block: TableBlock = serde_json::from_value(json!({
            "id": 0,
            "rows": [
                { "id": 0, "cells": [
                    { "id": 0, "rowSpan": 2, "blocks": [para()] },
                    { "id": 1, "blocks": [para()] },
                ] },
                { "id": 1, "height": 60.0, "heightRule": "exact", "cells": [
                    { "id": 2, "blocks": [para()] },
                ] },
            ],
            "columnWidths": [100.0, 100.0],
        }))
        .unwrap();
        let measure: TableExtent = serde_json::from_value(json!({
            "columnWidths": [100.0, 100.0],
            "totalWidth": 200.0,
            "totalHeight": 80.0,
            "rows": [
                { "height": 20.0, "cells": [
                    measured_cell(vec![para_measure(4)]),
                    measured_cell(vec![para_measure(1)]),
                ] },
                { "height": 60.0, "cells": [measured_cell(vec![para_measure(1)])] },
            ],
        }))
        .unwrap();
        assert!(block.rows[1].is_exact_height());
        assert!(!block.rows[0].is_exact_height());
        let info = build_table_row_break_info(&block, &measure);
        assert_eq!(info.break_offsets[0], vec![20.0]);
        assert_eq!(info.break_offsets[1], vec![60.0]);
        assert_eq!(snap_row_break(&info, 1, 0.0, 30.0), 0.0);
    }

    #[test]
    fn exact_row_taller_than_the_page_still_progresses() {
        // A 500px exact row on a 100px content column cannot move whole to a
        // fresh page and fit; pagination must place it with overflow rather
        // than loop forever pushing it forward.
        let table_block = json!({
            "kind": "table", "id": 0, "columnWidths": [100.0],
            "rows": [{
                "id": 0, "height": 500.0, "heightRule": "exact",
                "cells": [{ "id": 1, "blocks": [
                    { "kind": "paragraph", "id": 2, "runs": [] }
                ] }],
            }],
        });
        let table_measure = json!({
            "kind": "table", "columnWidths": [100.0],
            "totalWidth": 100.0, "totalHeight": 500.0,
            "rows": [{ "height": 500.0, "cells": [{
                "blocks": [{
                    "kind": "paragraph", "lines": [
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                          "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": 20.0 }
                    ], "totalHeight": 20.0,
                }],
                "width": 100.0, "height": 20.0,
            }] }],
        });
        let input = serde_json::json!({
            "measured": [{ "block": table_block, "measure": table_measure }],
            "options": {
                "pageSize": { "w": 400.0, "h": 100.0 },
                "margins": { "top": 0.0, "right": 0.0, "bottom": 0.0, "left": 0.0 },
            },
        })
        .to_string();
        let layout = crate::compute_layout(&input).expect("tall exact row lays out");
        assert!(!layout.pages.is_empty(), "makes progress");
        let table_fragments: Vec<_> = layout
            .pages
            .iter()
            .flat_map(|page| page.fragments.iter())
            .filter_map(|fragment| match fragment {
                crate::types::Fragment::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(table_fragments.len(), 1);
        assert_eq!(
            (table_fragments[0].row_start, table_fragments[0].row_end),
            (0, 1)
        );
        assert_eq!(table_fragments[0].clip_bottom, None);
    }

    #[test]
    fn exact_row_moves_whole_to_the_next_page_when_it_fits_there() {
        // Paragraph (80px) leaves 20px on a 100px page; the following 60px
        // exact row fits on an empty page, so it must move whole instead of
        // leaving a 20px slice behind.
        let para_block = json!({
            "kind": "paragraph", "id": 10, "runs": [],
        });
        let para_measure = json!({
            "kind": "paragraph",
            "lines": [{
                "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                "width": 0.0, "ascent": 64.0, "descent": 16.0, "lineHeight": 80.0,
            }],
            "totalHeight": 80.0,
        });
        let table_block = json!({
            "kind": "table", "id": 20, "columnWidths": [100.0],
            "rows": [{
                "id": 21, "height": 60.0, "heightRule": "exact",
                "cells": [{ "id": 22, "blocks": [
                    { "kind": "paragraph", "id": 23, "runs": [] }
                ] }],
            }],
        });
        let table_measure = json!({
            "kind": "table", "columnWidths": [100.0],
            "totalWidth": 100.0, "totalHeight": 60.0,
            "rows": [{ "height": 60.0, "cells": [{
                "blocks": [{
                    "kind": "paragraph",
                    "lines": [
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                          "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": 20.0 },
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                          "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": 20.0 },
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                          "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": 20.0 },
                    ],
                    "totalHeight": 60.0,
                }],
                "width": 100.0, "height": 60.0,
            }] }],
        });
        let input = serde_json::json!({
            "measured": [
                { "block": para_block, "measure": para_measure },
                { "block": table_block, "measure": table_measure },
            ],
            "options": {
                "pageSize": { "w": 400.0, "h": 100.0 },
                "margins": { "top": 0.0, "right": 0.0, "bottom": 0.0, "left": 0.0 },
            },
        })
        .to_string();
        let layout = crate::compute_layout(&input).expect("exact push-forward lays out");
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(layout.pages[0].fragments.len(), 1);
        let table_fragments: Vec<_> = layout
            .pages
            .iter()
            .flat_map(|page| page.fragments.iter())
            .filter_map(|fragment| match fragment {
                crate::types::Fragment::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(table_fragments.len(), 1);
        assert_eq!(
            (table_fragments[0].row_start, table_fragments[0].row_end),
            (0, 1)
        );
        assert_eq!(table_fragments[0].clip_top, None);
        assert_eq!(table_fragments[0].clip_bottom, None);
    }
}
