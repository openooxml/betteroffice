use crate::paragraph_spacing::{get_spacing_after, get_spacing_before};
use crate::types::{BlockExtent, ColumnLayout, LayoutBlock, MeasuredBlock};

pub trait ColumnBalancePaginator {
    fn columns(&self) -> ColumnLayout;
    fn pen_y(&mut self) -> f64;
    fn content_limit(&mut self) -> f64;
    fn set_content_limit(&mut self, value: f64);
}

struct SectionBalance {
    total_height: f64,
    legal_breaks: Vec<f64>,
}

fn get_balanced_section_height(
    measured: &[MeasuredBlock],
    start: usize,
    end: usize,
) -> Option<SectionBalance> {
    let mut total_height = 0.0_f64;
    let mut has_content = false;
    let mut legal_breaks = Vec::new();

    for mb in &measured[start..end] {
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
            total_height += (measure.total_height - measured_line_height).max(0.0);
            total_height += get_spacing_after(block);
            if block.attrs.as_ref().and_then(|attrs| attrs.keep_next) != Some(true) {
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
        })
    } else {
        None
    }
}

fn balance_current_column_region<P: ColumnBalancePaginator>(
    paginator: &mut P,
    total_content_height: f64,
    legal_breaks: &[f64],
) {
    let columns = paginator.columns();
    if columns.count <= 1.0 || !total_content_height.is_finite() || total_content_height <= 0.0 {
        return;
    }

    let column_region_top = paginator.pen_y();
    let max_region_height = paginator.content_limit() - column_region_top;
    if max_region_height <= 0.0 || total_content_height > max_region_height * columns.count {
        return;
    }

    let ideal = (total_content_height / columns.count).ceil();
    let balanced_height = legal_breaks
        .iter()
        .copied()
        .find(|boundary| *boundary >= ideal)
        .unwrap_or(ideal);
    if balanced_height <= 0.0 || balanced_height >= max_region_height {
        return;
    }

    paginator.set_content_limit(column_region_top + balanced_height);
}

pub fn balance_terminal_continuous_text_columns<P: ColumnBalancePaginator>(
    measured: &[MeasuredBlock],
    paginator: &mut P,
    start: usize,
    end: usize,
) {
    if let Some(balance) = get_balanced_section_height(measured, start, end) {
        balance_current_column_region(paginator, balance.total_height, &balance.legal_breaks);
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
                },
                pen_y,
                content_limit,
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

    #[test]
    fn balances_terminal_two_column_text_section() {
        let measured = vec![section_break(), text_paragraph("two-column", 6, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 130.0, 450.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![190.0]);
        assert_eq!(paginator.content_limit, 190.0);
    }

    // Five 20px lines have an ideal height of 50; the next legal line/widow
    // boundary is 60, so balancing never cuts a line in half.
    #[test]
    fn balanced_height_snaps_to_a_legal_line_boundary() {
        let measured = vec![text_paragraph("odd-lines", 5, 20.0)];
        let mut paginator = MockPaginator::new(2.0, 100.0, 400.0);
        balance_terminal_continuous_text_columns(&measured, &mut paginator, 0, measured.len());
        assert_eq!(paginator.set_calls, vec![160.0]);
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
        assert_eq!(paginator.set_calls, vec![156.0]);
    }

    #[test]
    fn empty_paragraph_spacing_collapses_unless_explicit() {
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
        assert_eq!(get_spacing_after(&empty), 0.0);

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
        assert_eq!(get_spacing_after(&explicit), 0.0);

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
        assert_eq!(paginator.set_calls, vec![140.0]);
    }
}
