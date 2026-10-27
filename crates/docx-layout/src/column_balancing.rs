//! Column balancing for a document's final continuous multi-column section.
//!
//! Word leaves the last continuous section's columns ragged unless it can even
//! them out. Balancing works by *shortening the region*: the content limit
//! drops so each column fills to roughly the same depth instead of the first
//! column running to the page bottom.
//!
//! The target is `ceil(total / columnCount)`, then snapped forward to the first
//! legal break at or past it, so a column boundary never lands mid-line. A break
//! between the exact `total / columnCount` and that ceiling wins instead when
//! all content still packs into the columns at it, so fractional lines split at
//! the half-way break. That packing model is only trusted where placement
//! follows it exactly: no table repeats a header row, no paragraph or table row
//! keeps with the next one, and no spacing is deferred into the region's top.
//! The exact comparison and the lowered limit carry `BOUNDARY_SLACK_PX`, so a
//! break that sits exactly on the target, or a column filled exactly to the
//! limit, is not lost to float drift. Legal breaks are:
//!
//! - inside a paragraph, after a line that is neither the first nor within two
//!   lines of the last, and only when the paragraph does not set `w:keepLines`;
//! - at a paragraph's end, unless it sets `w:keepNext`;
//! - after every table row, and after an image or text box.
//!
//! Balancing is abandoned — leaving the region at full height — when the
//! section has no content, when the range holds any block kind other than those
//! above (section breaks are simply skipped), when there is a single column,
//! when the content cannot fit the region even across all columns, or when the
//! snapped height is not actually shorter than the region.

use crate::hooks::row_keep_chains;
use crate::measure_blocks::measured_section_break_marks;
use crate::paragraph_spacing::{get_spacing_after, get_spacing_before, is_empty_paragraph};
use crate::types::{BlockExtent, ColumnLayout, LayoutBlock, MeasuredBlock};

/// Tolerance for float drift between stacked heights and the placed pen.
const BOUNDARY_SLACK_PX: f64 = 1e-6;

/// Page-flow operations required by column balancing.
pub trait ColumnBalancePaginator {
    fn columns(&self) -> ColumnLayout;
    fn pen_y(&mut self) -> f64;
    fn content_limit(&mut self) -> f64;
    fn set_content_limit(&mut self, value: f64);
    fn has_float_band_in_region(&mut self, _top: f64, _bottom: f64) -> bool {
        false
    }
    fn deferred_spacing(&mut self) -> f64 {
        0.0
    }
}

/// Stacked height of the balanced range and the offsets a column may end at,
/// both measured from the top of the range.
struct SectionBalance {
    total_height: f64,
    legal_breaks: Vec<f64>,
    /// No repeated table header or keep-with-next paragraph or row chain, so
    /// placement fills columns exactly as `content_packs` models them.
    packs_exactly: bool,
}

/// Stacks the range and collects its legal breaks, or `None` when the range
/// holds nothing to balance or a block kind balancing cannot reason about.
fn get_balanced_section_height(
    measured: &[MeasuredBlock],
    start: usize,
    end: usize,
    section_marks: &[bool],
) -> Option<SectionBalance> {
    let mut total_height = 0.0_f64;
    let mut has_content = false;
    let mut legal_breaks = Vec::new();
    let mut packs_exactly = true;

    for (index, mb) in measured.iter().enumerate().take(end).skip(start) {
        if section_marks[index] {
            continue;
        }
        if let (LayoutBlock::Paragraph(block), BlockExtent::Paragraph(measure)) =
            (&mb.block, &mb.measure)
        {
            total_height += get_spacing_before(block);
            let mut measured_line_height = 0.0_f64;
            for (line_index, line) in measure.lines.iter().enumerate() {
                let line_height = line.line_height + line.float_skip_before.unwrap_or(0.0);
                measured_line_height += line_height;
                total_height += line_height;
                let lines_after = measure.lines.len() - line_index - 1;
                let legal_inside = block.attrs.as_ref().and_then(|attrs| attrs.keep_lines)
                    != Some(true)
                    && line_index >= 1
                    && lines_after >= 2;
                if legal_inside {
                    legal_breaks.push(total_height);
                }
            }
            if !is_empty_paragraph(block) {
                total_height += (measure.total_height - measured_line_height).max(0.0);
            }
            total_height += get_spacing_after(block);
            if block.attrs.as_ref().and_then(|attrs| attrs.keep_next) == Some(true) {
                packs_exactly = false;
            } else {
                legal_breaks.push(total_height);
            }
            has_content = has_content || !measure.lines.is_empty();
            continue;
        }

        if matches!(mb.block, LayoutBlock::SectionBreak(_)) {
            continue;
        }

        match &mb.measure {
            BlockExtent::Table(table) => {
                if let LayoutBlock::Table(block) = &mb.block {
                    let repeats_header = block
                        .rows
                        .first()
                        .is_some_and(|row| row.is_header == Some(true));
                    let keeps_rows = row_keep_chains(block, table).iter().any(Option::is_some);
                    packs_exactly &= !repeats_header && !keeps_rows;
                }
                for row in &table.rows {
                    total_height += row.height;
                    legal_breaks.push(total_height);
                }
                has_content = has_content || !table.rows.is_empty();
            }
            BlockExtent::Image(image) => {
                total_height += image.height;
                legal_breaks.push(total_height);
                has_content = true;
            }
            BlockExtent::TextBox(text_box) => {
                total_height += text_box.height;
                legal_breaks.push(total_height);
                has_content = true;
            }
            _ => return None,
        }
    }

    if has_content {
        Some(SectionBalance {
            total_height,
            legal_breaks,
            packs_exactly,
        })
    } else {
        None
    }
}

/// Whether `columns` columns of `height`, each filled to the furthest legal
/// break that fits, reach the end of the content.
fn content_packs(legal_breaks: &[f64], total_height: f64, columns: usize, height: f64) -> bool {
    let mut start = 0.0_f64;
    for _ in 0..columns {
        let limit = start + height + BOUNDARY_SLACK_PX;
        if total_height <= limit {
            return true;
        }
        match legal_breaks
            .iter()
            .rev()
            .copied()
            .find(|boundary| *boundary > start && *boundary <= limit)
        {
            Some(boundary) => start = boundary,
            None => return false,
        }
    }
    false
}

/// Lowers the current region's content limit to the balanced depth, or leaves
/// it alone when any of the module's bail-out conditions holds.
fn balance_current_column_region<P: ColumnBalancePaginator>(
    paginator: &mut P,
    total_content_height: f64,
    legal_breaks: &[f64],
    packs_exactly: bool,
) {
    let columns = paginator.columns();
    if columns.count <= 1.0 || !total_content_height.is_finite() || total_content_height <= 0.0 {
        return;
    }

    let column_region_top = paginator.pen_y();
    let content_limit = paginator.content_limit();
    if paginator.has_float_band_in_region(column_region_top, content_limit) {
        return;
    }
    let max_region_height = content_limit - column_region_top;
    if max_region_height <= 0.0 || total_content_height > max_region_height * columns.count {
        return;
    }

    let ideal = total_content_height / columns.count;
    let ceiled = ideal.ceil();
    let packs_exactly = packs_exactly && paginator.deferred_spacing() == 0.0;
    let half_way = legal_breaks
        .iter()
        .copied()
        .filter(|boundary| *boundary >= ideal - BOUNDARY_SLACK_PX && *boundary < ceiled)
        .find(|height| {
            packs_exactly
                && content_packs(
                    legal_breaks,
                    total_content_height,
                    columns.count as usize,
                    *height,
                )
        });
    let balanced_height = half_way.unwrap_or_else(|| {
        legal_breaks
            .iter()
            .copied()
            .find(|boundary| *boundary >= ceiled)
            .unwrap_or(ceiled)
    });
    if balanced_height <= 0.0 || balanced_height >= max_region_height {
        return;
    }

    paginator.set_content_limit(
        (column_region_top + balanced_height + BOUNDARY_SLACK_PX).min(content_limit),
    );
}

/// Balances the measured blocks in `start..end` across the current region's
/// columns. Called with the range that follows the terminal continuous break.
pub fn balance_terminal_continuous_text_columns<P: ColumnBalancePaginator>(
    measured: &[MeasuredBlock],
    paginator: &mut P,
    start: usize,
    end: usize,
) {
    let marks = measured_section_break_marks(measured, None, None);
    balance_terminal_continuous_text_columns_with_marks(measured, paginator, start, end, &marks);
}

pub(crate) fn balance_terminal_continuous_text_columns_with_marks<P: ColumnBalancePaginator>(
    measured: &[MeasuredBlock],
    paginator: &mut P,
    start: usize,
    end: usize,
    section_marks: &[bool],
) {
    if let Some(balance) = get_balanced_section_height(measured, start, end, section_marks) {
        balance_current_column_region(
            paginator,
            balance.total_height,
            &balance.legal_breaks,
            balance.packs_exactly,
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{
        BlockId, ParagraphAttrs, ParagraphBlock, ParagraphExtent, ParagraphSpacing, Run,
        RunFormatting, SectionBreakBlock, SpacingExplicit, TextRun, TypesetRow,
    };

    struct MockPaginator {
        columns: ColumnLayout,
        pen_y: f64,
        content_limit: f64,
        deferred_spacing: f64,
        set_calls: Vec<f64>,
    }

    impl MockPaginator {
        fn new(count: f64, pen_y: f64, content_limit: f64) -> Self {
            MockPaginator {
                columns: ColumnLayout {
                    count,
                    gap: 20.0,
                    equal_width: None,
                    separator: None,
                    columns: None,
                },
                pen_y,
                content_limit,
                deferred_spacing: 0.0,
                set_calls: Vec::new(),
            }
        }
    }

    impl ColumnBalancePaginator for MockPaginator {
        fn columns(&self) -> ColumnLayout {
            self.columns.clone()
        }
        fn pen_y(&mut self) -> f64 {
            self.pen_y
        }
        fn content_limit(&mut self) -> f64 {
            self.content_limit
        }
        fn set_content_limit(&mut self, value: f64) {
            self.content_limit = value;
            self.set_calls.push(value);
        }
        fn deferred_spacing(&mut self) -> f64 {
            self.deferred_spacing
        }
    }

    fn text_run(text: &str) -> Run {
        Run::Text(TextRun {
            fmt: RunFormatting::default(),
            text: text.to_string(),
            pm_start: None,
            pm_end: None,
            inline_sdt_widget: None,
        })
    }

    fn para_block(runs: Vec<Run>, attrs: Option<ParagraphAttrs>) -> ParagraphBlock {
        ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            para_id: None,
            runs,
            attrs,
            pm_start: None,
            pm_end: None,
        }
    }

    fn text_paragraph(text: &str, line_count: usize, line_height: f64) -> MeasuredBlock {
        MeasuredBlock {
            block: LayoutBlock::Paragraph(para_block(vec![text_run(text)], None)),
            measure: BlockExtent::Paragraph(ParagraphExtent {
                lines: (0..line_count)
                    .map(|_| TypesetRow {
                        line_height,
                        marker_tab_offset: None,
                        ..Default::default()
                    })
                    .collect(),
                total_height: line_count as f64 * line_height,
            }),
        }
    }

    fn section_break() -> MeasuredBlock {
        MeasuredBlock {
            block: LayoutBlock::SectionBreak(SectionBreakBlock {
                sdt_groups: None,
                id: BlockId::Num(0.0),
                break_type: None,
                page_size: None,
                orientation: None,
                margins: None,
                columns: None,
            }),
            measure: BlockExtent::SectionBreak,
        }
    }

    fn other_block() -> MeasuredBlock {
        MeasuredBlock {
            block: LayoutBlock::Unsupported,
            measure: BlockExtent::Unsupported,
        }
    }

    // page 500x500 / margins 50, intro paragraph of 80 above the region, then
    // a 6-line x 20 paragraph balanced over two columns → the region limit
    // drops from 450 to 130 + 120 / 2 = 190 (3 lines per column).
    #[test]
    fn balances_terminal_two_column_text_section() {
        let measured = vec![section_break(), text_paragraph("two-column", 6, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 130.0, 450.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![190.0 + BOUNDARY_SLACK_PX]);
        assert_eq!(paginator.content_limit, 190.0 + BOUNDARY_SLACK_PX);
    }

    // Five 20px lines have an ideal height of 50; the next legal line/widow
    // boundary is 60, so balancing never cuts a line in half.
    #[test]
    fn balanced_height_snaps_to_a_legal_line_boundary() {
        let measured = vec![text_paragraph("odd-lines", 5, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![160.0 + BOUNDARY_SLACK_PX]);
    }

    // Eight 36.8px one-line paragraphs and the empty paragraph carrying the
    // section break split 4/4: the half-way break is legal even though the
    // stacked height is not a whole pixel, and the limit still admits the
    // fourth line once the pen has drifted past the exact sum.
    #[test]
    fn non_integer_lines_balance_at_the_half_way_break() {
        let line = 36.8;
        let mut measured: Vec<MeasuredBlock> =
            (0..8).map(|_| text_paragraph("item", 1, line)).collect();
        measured.push(MeasuredBlock {
            block: LayoutBlock::Paragraph(para_block(vec![], None)),
            measure: BlockExtent::Paragraph(ParagraphExtent {
                lines: Vec::new(),
                total_height: 0.0,
            }),
        });
        measured.push(section_break());
        let marks = vec![false; measured.len()];
        let mut paginator = MockPaginator::new(2.0, 100.0, 900.0);
        balance_terminal_continuous_text_columns_with_marks(
            &measured,
            &mut paginator,
            0,
            measured.len() - 1,
            &marks,
        );

        let limit = paginator.content_limit;
        assert!(limit < 100.0 + 4.5 * line, "limit {limit}");
        let mut pen = 100.0;
        let mut placed = 0;
        while placed < 8 && limit - pen >= line {
            pen += line;
            placed += 1;
        }
        assert_eq!(placed, 4);
    }

    // Total 220 over three columns: the 73.6 break nearest 220 / 3 leaves 146.4
    // for two columns that can only take [40] and [40], so 66.4 would spill.
    // The smallest break that packs all content is 113.6.
    #[test]
    fn three_columns_take_the_smallest_break_that_packs_all_content() {
        let measured: Vec<MeasuredBlock> = [36.8, 36.8, 40.0, 40.0, 66.4]
            .iter()
            .map(|height| text_paragraph("item", 1, *height))
            .collect();
        let mut paginator = MockPaginator::new(3.0, 100.0, 300.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(
            paginator.set_calls,
            vec![100.0 + (36.8 + 36.8 + 40.0) + BOUNDARY_SLACK_PX]
        );
    }

    fn attr_paragraph(lines: &[f64], total_height: f64, attrs: ParagraphAttrs) -> MeasuredBlock {
        MeasuredBlock {
            block: LayoutBlock::Paragraph(para_block(vec![text_run("x")], Some(attrs))),
            measure: BlockExtent::Paragraph(ParagraphExtent {
                lines: lines
                    .iter()
                    .map(|line_height| TypesetRow {
                        line_height: *line_height,
                        ..Default::default()
                    })
                    .collect(),
                total_height,
            }),
        }
    }

    fn keep_next_attrs() -> ParagraphAttrs {
        ParagraphAttrs {
            keep_next: Some(true),
            widow_control: Some(false),
            ..Default::default()
        }
    }

    fn table(row_heights: &[f64], header_rows: usize) -> MeasuredBlock {
        keep_next_table(row_heights, header_rows, &[])
    }

    fn keep_next_table(
        row_heights: &[f64],
        header_rows: usize,
        keep_next_rows: &[usize],
    ) -> MeasuredBlock {
        let keep_next_cell = serde_json::json!({
            "id": 0,
            "blocks": [LayoutBlock::Paragraph(para_block(
                vec![text_run("x")],
                Some(keep_next_attrs()),
            ))],
        });
        let rows: Vec<serde_json::Value> = (0..row_heights.len())
            .map(|index| {
                let cells = if keep_next_rows.contains(&index) {
                    vec![keep_next_cell.clone()]
                } else {
                    Vec::new()
                };
                serde_json::json!({ "id": index, "cells": cells, "isHeader": index < header_rows })
            })
            .collect();
        let extents: Vec<serde_json::Value> = row_heights
            .iter()
            .map(|height| serde_json::json!({ "cells": [], "height": height }))
            .collect();
        MeasuredBlock {
            block: LayoutBlock::Table(
                serde_json::from_value(serde_json::json!({ "id": 0, "rows": rows })).unwrap(),
            ),
            measure: BlockExtent::Table(
                serde_json::from_value(serde_json::json!({
                    "rows": extents,
                    "columnWidths": [],
                    "totalWidth": 0.0,
                    "totalHeight": row_heights.iter().sum::<f64>(),
                }))
                .unwrap(),
            ),
        }
    }

    // A continuation fragment repeats the header row, which the half-way 4/4
    // split does not leave room for; the ceiled target keeps all rows on the
    // page.
    #[test]
    fn repeated_table_headers_keep_the_ceiled_target() {
        let measured = vec![table(&[36.125; 8], 1)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 500.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(
            paginator.set_calls,
            vec![100.0 + 5.0 * 36.125 + BOUNDARY_SLACK_PX]
        );
    }

    // The half-way 100.5 falls after line 2 of the keepNext paragraph, but
    // placement moves that paragraph with its follower whole: column 0 would
    // end at 60.5 and the last paragraph spill to a second page.
    #[test]
    fn a_keep_next_paragraph_keeps_main_target() {
        let keep_next_lines = ParagraphAttrs {
            keep_next: Some(true),
            keep_lines: Some(false),
            ..Default::default()
        };
        let measured = vec![
            text_paragraph("a", 1, 60.5),
            attr_paragraph(&[20.0; 4], 80.0, keep_next_lines),
            text_paragraph("c", 1, 20.0),
            text_paragraph("d", 1, 40.0),
        ];
        let mut paginator = MockPaginator::new(2.0, 100.0, 500.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![100.0 + 160.5 + BOUNDARY_SLACK_PX]);
    }

    // Row 1 keeps with row 2, so placement cannot break after row 1 at the
    // half-way 100.5: rows 1 and 2 move to column 1 and row 3 spills.
    #[test]
    fn a_table_row_keep_chain_keeps_main_target() {
        let measured = vec![keep_next_table(&[60.5, 40.0, 40.0, 60.0], 0, &[1])];
        let mut paginator = MockPaginator::new(2.0, 100.0, 500.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![100.0 + 140.5 + BOUNDARY_SLACK_PX]);
    }

    // The 12px spacing deferred from the paragraph above the continuous break
    // is charged to column 0, which then takes only three 36.8px lines at the
    // half-way 147.2 and the eighth spills; main's 184 keeps 4/4.
    #[test]
    fn spacing_deferred_into_the_region_keeps_main_target() {
        let measured: Vec<MeasuredBlock> =
            (0..8).map(|_| text_paragraph("item", 1, 36.8)).collect();
        let mut paginator = MockPaginator::new(2.0, 150.0, 900.0);
        paginator.deferred_spacing = 12.0;
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        let five_lines = (0..5).fold(0.0, |height, _| height + 36.8);
        assert_eq!(
            paginator.set_calls,
            vec![150.0 + five_lines + BOUNDARY_SLACK_PX]
        );
    }

    #[test]
    fn content_without_legal_breaks_keeps_the_ceiled_target() {
        let measured = vec![attr_paragraph(&[30.5, 10.5, 40.5], 81.5, keep_next_attrs())];
        let mut paginator = MockPaginator::new(2.0, 100.0, 300.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![141.0 + BOUNDARY_SLACK_PX]);
    }

    // Stacked spacing makes the greedy model reject 100, although placement
    // collapses it and fits; a whole-pixel target is main's choice.
    #[test]
    fn a_whole_pixel_target_is_not_second_guessed_by_the_packing_model() {
        let spaced = |before: f64, after: f64| ParagraphAttrs {
            spacing: Some(ParagraphSpacing {
                before: Some(before),
                after: Some(after),
                ..Default::default()
            }),
            ..Default::default()
        };
        let measured = vec![
            text_paragraph("a", 1, 50.0),
            text_paragraph("b", 1, 50.0),
            attr_paragraph(&[20.0], 40.0, spaced(0.0, 20.0)),
            attr_paragraph(&[20.0], 40.0, spaced(20.0, 0.0)),
            text_paragraph("e", 1, 80.0),
        ];
        let mut paginator = MockPaginator::new(3.0, 100.0, 500.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![200.0 + BOUNDARY_SLACK_PX]);
    }

    #[test]
    fn a_trailing_keep_next_paragraph_still_balances() {
        let measured = vec![
            text_paragraph("a", 1, 20.0),
            text_paragraph("b", 1, 20.0),
            attr_paragraph(&[30.0, 30.0], 60.0, keep_next_attrs()),
        ];
        let mut paginator = MockPaginator::new(3.0, 100.0, 300.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![140.0 + BOUNDARY_SLACK_PX]);
    }

    #[test]
    fn spacing_before_and_after_count_toward_the_balanced_height() {
        let block = para_block(
            vec![text_run("x")],
            Some(ParagraphAttrs {
                spacing: Some(ParagraphSpacing {
                    before: Some(10.0),
                    after: Some(6.0),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        assert_eq!(get_spacing_before(&block), 10.0);
        assert_eq!(get_spacing_after(&block), 6.0);

        // A two-line paragraph has no legal internal widow boundary, so the
        // whole 56px paragraph stays in the first column.
        let measured = vec![MeasuredBlock {
            block: LayoutBlock::Paragraph(block),
            measure: BlockExtent::Paragraph(ParagraphExtent {
                lines: vec![TypesetRow::default(), TypesetRow::default()],
                total_height: 40.0,
            }),
        }];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![156.0 + BOUNDARY_SLACK_PX]);
    }

    #[test]
    fn empty_paragraph_after_spacing_is_counted_once_when_measured() {
        let block = para_block(
            vec![],
            Some(ParagraphAttrs {
                spacing: Some(ParagraphSpacing {
                    after: Some(12.0),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let measured = vec![MeasuredBlock {
            block: LayoutBlock::Paragraph(block),
            measure: BlockExtent::Paragraph(ParagraphExtent {
                lines: vec![TypesetRow {
                    line_height: 16.0,
                    marker_tab_offset: None,
                    ..Default::default()
                }],
                total_height: 28.0,
            }),
        }];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![128.0 + BOUNDARY_SLACK_PX]);
    }

    #[test]
    fn empty_paragraph_after_spacing_preserves_inherited_values() {
        let empty = para_block(
            vec![text_run("")],
            Some(ParagraphAttrs {
                spacing: Some(ParagraphSpacing {
                    before: Some(12.0),
                    after: Some(12.0),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        assert_eq!(get_spacing_before(&empty), 0.0);
        assert_eq!(get_spacing_after(&empty), 12.0);

        let explicit = para_block(
            vec![],
            Some(ParagraphAttrs {
                spacing: Some(ParagraphSpacing {
                    before: Some(12.0),
                    after: Some(12.0),
                    ..Default::default()
                }),
                spacing_explicit: Some(SpacingExplicit {
                    before: Some(true),
                    after: None,
                }),
                ..Default::default()
            }),
        );
        assert_eq!(get_spacing_before(&explicit), 12.0);
        assert_eq!(get_spacing_after(&explicit), 12.0);

        // multi-run paragraphs are never "empty"
        let multi = para_block(
            vec![text_run(""), text_run("")],
            Some(ParagraphAttrs {
                spacing: Some(ParagraphSpacing {
                    before: Some(7.0),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        assert_eq!(get_spacing_before(&multi), 7.0);
    }

    #[test]
    fn single_column_region_is_not_balanced() {
        let measured = vec![text_paragraph("solo", 4, 20.0)];
        let mut paginator = MockPaginator::new(1.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn non_text_block_in_range_disables_balancing() {
        let measured = vec![text_paragraph("text", 3, 20.0), other_block()];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn paragraph_block_with_non_paragraph_measure_disables_balancing() {
        let measured = vec![MeasuredBlock {
            block: LayoutBlock::Paragraph(para_block(vec![], None)),
            measure: BlockExtent::Unsupported,
        }];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn section_with_no_text_lines_is_not_balanced() {
        let measured = vec![section_break(), text_paragraph("", 0, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn content_taller_than_the_region_capacity_is_not_balanced() {
        // region 100..400 (300 tall), 2 columns → capacity 600 < total 700
        let measured = vec![text_paragraph("tall", 35, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn balanced_height_at_or_above_region_height_is_not_applied() {
        // total 500 over 2 columns → 250, but region is only 250 tall →
        // balancedHeight >= maxRegionHeight → no-op
        let measured = vec![text_paragraph("full", 25, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 150.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn exhausted_region_is_not_balanced() {
        // penY at contentLimit → maxRegionHeight = 0 → no-op
        let measured = vec![text_paragraph("late", 2, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 400.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert!(paginator.set_calls.is_empty());
    }

    #[test]
    fn start_end_slice_only_considers_the_terminal_section() {
        let measured = vec![other_block(), text_paragraph("tail", 4, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 1, measured.len());
        assert_eq!(paginator.set_calls, vec![140.0 + BOUNDARY_SLACK_PX]);
    }
}
