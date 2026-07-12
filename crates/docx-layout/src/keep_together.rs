use std::collections::{BTreeMap, BTreeSet};

use crate::paragraph_spacing::{get_spacing_after, get_spacing_before};
use crate::types::{BlockExtent, LayoutBlock, MeasuredBlock};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeepWithNextGroup {
    /// Index of the run's leading paragraph.
    pub head_index: usize,
    /// Index of the run's final keep-with-next paragraph.
    pub tail_index: usize,
    /// Every block index that belongs to the run, in order.
    pub members: Vec<usize>,
    /// Index of the following flow block whose first unbreakable slice is the
    /// keep witness, or `None` at a forced/section break or EOF.
    pub follower: Option<usize>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct KeepWithNextScan {
    /// Groups keyed by their leading block index.
    pub groups_by_head: BTreeMap<usize, KeepWithNextGroup>,
    /// Block indices that belong to a group but are not its head.
    pub interior_members: BTreeSet<usize>,
}

// true only for a paragraph block carrying a truthy keepNext flag
fn is_bound_paragraph(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(p) => p.attrs.as_ref().and_then(|a| a.keep_next).unwrap_or(false),
        _ => false,
    }
}

/// Group every maximal run of consecutive keep-with-next paragraphs.
///
/// A run grows while the next block is another keep-with-next paragraph; it
/// ends at a break block, a non-paragraph block, a paragraph without keepNext,
/// or the end of the list. When the terminator is a plain paragraph it becomes
/// the run's follower, since the run must land on the follower's page.
pub fn analyze_keep_with_next(measured: &[MeasuredBlock]) -> KeepWithNextScan {
    let mut groups_by_head: BTreeMap<usize, KeepWithNextGroup> = BTreeMap::new();
    let mut interior_members: BTreeSet<usize> = BTreeSet::new();

    let mut cursor = 0usize;
    while cursor < measured.len() {
        if !is_bound_paragraph(&measured[cursor].block) {
            cursor += 1;
            continue;
        }

        let mut members: Vec<usize> = vec![cursor];
        let mut tail_index = cursor;
        let mut probe = cursor + 1;
        while probe < measured.len() && is_bound_paragraph(&measured[probe].block) {
            members.push(probe);
            tail_index = probe;
            probe += 1;
        }

        // A keep chain binds to the first unbreakable slice of any following
        // supported flow object. Forced/section breaks terminate it.
        let after_tail = tail_index + 1;
        let follower = if after_tail < measured.len()
            && matches!(
                measured[after_tail].block,
                LayoutBlock::Paragraph(_)
                    | LayoutBlock::Table(_)
                    | LayoutBlock::Image(_)
                    | LayoutBlock::TextBox(_)
            ) {
            Some(after_tail)
        } else {
            None
        };

        for member in members.iter().skip(1).copied() {
            interior_members.insert(member);
        }
        groups_by_head.insert(
            cursor,
            KeepWithNextGroup {
                head_index: cursor,
                tail_index,
                members,
                follower,
            },
        );

        cursor = tail_index + 1;
    }

    KeepWithNextScan {
        groups_by_head,
        interior_members,
    }
}

pub fn measure_keep_with_next_group(group: &KeepWithNextGroup, measured: &[MeasuredBlock]) -> f64 {
    // follower's witness line first: zero when there is no follower, or when
    // it is not a laid-out paragraph
    let follower_measure = group
        .follower
        .and_then(|index| measured.get(index))
        .map(|mb| &mb.measure);
    let witness_line = match follower_measure {
        Some(BlockExtent::Paragraph(p)) if !p.lines.is_empty() => p.lines[0].line_height,
        Some(BlockExtent::Table(t)) => t.rows.first().map_or(0.0, |row| row.height),
        Some(BlockExtent::Image(image)) => image.height,
        Some(BlockExtent::TextBox(text_box)) => text_box.height,
        _ => 0.0,
    };

    let mut budget = witness_line;
    for &index in &group.members {
        let MeasuredBlock { block, measure } = &measured[index];
        let (LayoutBlock::Paragraph(block), BlockExtent::Paragraph(measure)) = (block, measure)
        else {
            continue;
        };
        budget += get_spacing_before(block) + measure.total_height + get_spacing_after(block);
    }

    budget
}

/// Whether a paragraph forbids splitting its own lines across a page (keepLines).
pub fn paragraph_keeps_lines(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(p) => p.attrs.as_ref().and_then(|a| a.keep_lines) == Some(true),
        _ => false,
    }
}

/// Whether a paragraph must begin on a fresh page (pageBreakBefore).
pub fn paragraph_breaks_before(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(p) => {
            p.attrs.as_ref().and_then(|a| a.page_break_before) == Some(true)
        }
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::break_policy::{KeepWithNextFit, keep_with_next_group_must_advance};
    use crate::types::{
        BlockId, ParagraphAttrs, ParagraphBlock, ParagraphExtent, ParagraphSpacing, Run,
        RunFormatting, SpacingExplicit, TextRun, TypesetRow,
    };

    fn text_run(text: &str) -> Run {
        Run::Text(TextRun {
            fmt: RunFormatting::default(),
            text: text.to_string(),
            pm_start: None,
            pm_end: None,
            inline_sdt_widget: None,
        })
    }

    fn paragraph(runs: Vec<Run>, attrs: Option<ParagraphAttrs>) -> LayoutBlock {
        LayoutBlock::Paragraph(ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            para_id: None,
            runs,
            attrs,
            pm_start: None,
            pm_end: None,
        })
    }

    fn make_paragraph_block(text: &str, keep_next: bool) -> LayoutBlock {
        paragraph(
            vec![text_run(text)],
            Some(ParagraphAttrs {
                keep_next: if keep_next { Some(true) } else { None },
                ..Default::default()
            }),
        )
    }

    fn make_line(line_height: f64) -> TypesetRow {
        TypesetRow {
            line_height,
            ..Default::default()
        }
    }

    fn make_paragraph_measure(lines: Vec<TypesetRow>) -> BlockExtent {
        let total_height = lines.iter().map(|l| l.line_height).sum();
        BlockExtent::Paragraph(ParagraphExtent {
            lines,
            total_height,
        })
    }

    // empty keepNext paragraph whose spacing is style-inherited (spacingExplicit
    // unset) — placement drops this spacing, so the group estimate must too
    fn make_empty_spaced_paragraph(
        spacing: (f64, f64),
        spacing_explicit: Option<(bool, bool)>,
    ) -> LayoutBlock {
        paragraph(
            vec![text_run("")],
            Some(ParagraphAttrs {
                keep_next: Some(true),
                spacing: Some(ParagraphSpacing {
                    before: Some(spacing.0),
                    after: Some(spacing.1),
                    ..Default::default()
                }),
                spacing_explicit: spacing_explicit.map(|(before, after)| SpacingExplicit {
                    before: Some(before),
                    after: Some(after),
                }),
                ..Default::default()
            }),
        )
    }

    fn to_measured_blocks(
        blocks: Vec<LayoutBlock>,
        measures: Vec<BlockExtent>,
    ) -> Vec<MeasuredBlock> {
        assert_eq!(blocks.len(), measures.len());
        blocks
            .into_iter()
            .zip(measures)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect()
    }

    #[test]
    fn ignores_style_inherited_spacing_on_an_empty_member_like_placement_does() {
        let blocks = vec![
            make_paragraph_block("Heading", true),
            make_empty_spaced_paragraph((150.0, 150.0), None),
            make_paragraph_block("Follower", false),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(20.0)]),
            make_paragraph_measure(vec![]),
            make_paragraph_measure(vec![make_line(20.0)]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan.groups_by_head.get(&0);
        assert!(group.is_some());

        // heading line (20) + empty member (0, spacing suppressed) + follower first line (20)
        assert_eq!(
            measure_keep_with_next_group(group.unwrap(), &measured),
            40.0
        );
    }

    #[test]
    fn keeps_counting_explicit_spacing_on_an_empty_member() {
        let blocks = vec![
            make_paragraph_block("Heading", true),
            make_empty_spaced_paragraph((150.0, 150.0), Some((true, true))),
            make_paragraph_block("Follower", false),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(20.0)]),
            make_paragraph_measure(vec![]),
            make_paragraph_measure(vec![make_line(20.0)]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan.groups_by_head.get(&0);
        assert!(group.is_some());

        // direct formatting survives on empty paragraphs: 20 + (150 + 0 + 150) + 20
        assert_eq!(
            measure_keep_with_next_group(group.unwrap(), &measured),
            340.0
        );
    }

    #[test]
    fn does_not_advance_a_group_that_fits_once_inherited_empty_paragraph_spacing_is_dropped() {
        let blocks = vec![
            make_paragraph_block("Filler", false),
            make_paragraph_block("Heading", true),
            make_empty_spaced_paragraph((150.0, 150.0), None),
            make_paragraph_block("Follower", false),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(620.0)]),
            make_paragraph_measure(vec![make_line(20.0)]),
            make_paragraph_measure(vec![]),
            make_paragraph_measure(vec![make_line(20.0)]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan
            .groups_by_head
            .get(&1)
            .expect("group headed at block 1");
        assert_eq!(group.members, vec![1, 2]);
        assert_eq!(group.follower, Some(3));

        let group_height = measure_keep_with_next_group(group, &measured);
        assert_eq!(group_height, 40.0);

        // content height 864 (1056 - 2*96); the 620px filler leaves 244 available
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height,
            available_height: 244.0,
            page_content_height: 864.0,
            page_has_content: true,
        }));
    }
}
