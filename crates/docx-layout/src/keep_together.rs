//! Keep-with-next grouping (`w:keepNext`, ECMA-376 §17.3.1.15).
//!
//! A run of consecutive keepNext paragraphs must share a page with the *start*
//! of whatever follows it. [`analyze_keep_with_next`] walks the measured blocks
//! once and returns each run keyed by its head, plus the interior members, so
//! the placement walk can skip blocks a group already accounted for.
//!
//! [`measure_keep_with_next_group`] turns a group into the height the contract
//! actually demands: every member paragraph in full (spacing before, measured
//! height, spacing after) plus one witness slice of the follower — never the
//! follower in full, since the binding is only to where it begins. The witness
//! is a paragraph's shortest legally placeable leading slice
//! ([`paragraph_min_leading_slice`]), a table's initial header/body slice, the
//! height of an image or text box, and nothing at all for any other follower
//! kind.
//!
//! Spacing follows the shared paragraph-spacing helpers used by placement.

use std::collections::{BTreeMap, BTreeSet};

use crate::paragraph_spacing::{get_spacing_after, get_spacing_before, is_empty_paragraph};
use crate::table_row_break::{build_table_row_break_info, first_table_fragment_height};
use crate::types::{BlockExtent, LayoutBlock, MeasuredBlock, ParagraphBlock, ParagraphExtent};

/// Lines below which widow/orphan control forbids every internal break, since
/// each side of one needs two lines.
const WIDOW_CONTROL_MIN_LINES: usize = 2;

/// A maximal keep-with-next run and its follower.
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

/// The two indexes placement needs: groups by head block, and every non-head
/// member so the walk does not re-evaluate them. Both iterate in ascending
/// block order.
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
                    | LayoutBlock::Shape(_)
                    | LayoutBlock::Chart(_)
                    | LayoutBlock::TextBox(_)
            ) {
            Some(after_tail)
        } else {
            None
        };

        for k in 1..members.len() {
            interior_members.insert(members[k]);
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

/// Whether widow/orphan control governs this paragraph: `w:widowControl`
/// (ECMA-376 §17.3.1.44) defaults on, and one line has nothing to protect.
pub fn paragraph_widow_control(block: &ParagraphBlock, measure: &ParagraphExtent) -> bool {
    measure.lines.len() >= WIDOW_CONTROL_MIN_LINES
        && block
            .attrs
            .as_ref()
            .and_then(|attrs| attrs.widow_control)
            .unwrap_or(true)
}

/// Whether a paragraph refuses every internal page break: `w:keepLines`, or
/// widow/orphan control on a paragraph too short to leave two lines on both
/// sides of one.
pub fn paragraph_is_unbreakable(block: &ParagraphBlock, measure: &ParagraphExtent) -> bool {
    block
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.keep_lines)
        .unwrap_or(false)
        || (paragraph_widow_control(block, measure)
            && measure.lines.len() < 2 * WIDOW_CONTROL_MIN_LINES)
}

/// Height (px) of the shortest leading slice a paragraph may legally place:
/// every line when it refuses to split, two under widow/orphan control, else
/// one. Placement and the keepNext witness both read this so they cannot
/// disagree about how much of a follower must come along.
pub fn paragraph_min_leading_slice(block: &ParagraphBlock, measure: &ParagraphExtent) -> f64 {
    let lines = if paragraph_is_unbreakable(block, measure) {
        measure.lines.len()
    } else if paragraph_widow_control(block, measure) {
        WIDOW_CONTROL_MIN_LINES
    } else {
        1
    };
    measure
        .lines
        .iter()
        .take(lines)
        .map(|line| line.line_height + line.float_skip_before.unwrap_or(0.0))
        .sum()
}

/// Vertical space (px) the group needs for its keepNext contract to hold on a
/// single page: the members in full plus the follower's witness slice.
pub fn measure_keep_with_next_group(group: &KeepWithNextGroup, measured: &[MeasuredBlock]) -> f64 {
    // follower's witness slice first: zero when there is no follower, or when
    // it is not a laid-out paragraph
    let follower = group.follower.and_then(|index| measured.get(index));
    let witness_line = match follower.map(|mb| &mb.measure) {
        Some(BlockExtent::Paragraph(p)) if !p.lines.is_empty() => {
            match follower.map(|mb| &mb.block) {
                Some(LayoutBlock::Paragraph(block)) => paragraph_min_leading_slice(block, p),
                _ => p.lines[0].line_height,
            }
        }
        Some(BlockExtent::Table(table)) => match follower.map(|mb| &mb.block) {
            Some(LayoutBlock::Table(block)) => {
                first_table_fragment_height(block, table, &build_table_row_break_info(block, table))
            }
            _ => 0.0,
        },
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
        let height = if is_empty_paragraph(block) {
            measure
                .lines
                .iter()
                .map(|line| line.line_height + line.float_skip_before.unwrap_or(0.0))
                .sum()
        } else {
            measure.total_height
        };
        budget += get_spacing_before(block) + height + get_spacing_after(block);
    }

    budget
}

/// Whether a paragraph forbids splitting its own lines across a page (keepLines).
#[allow(dead_code)]
pub fn paragraph_keeps_lines(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(p) => p.attrs.as_ref().and_then(|a| a.keep_lines) == Some(true),
        _ => false,
    }
}

/// Whether a paragraph must begin on a fresh page, and whether that is
/// because a hard `w:br w:type="page"` run opens it rather than
/// `w:pageBreakBefore`.
pub fn paragraph_breaks_before_run(block: &LayoutBlock) -> (bool, bool) {
    match block {
        LayoutBlock::Paragraph(p) => {
            let attrs = p.attrs.as_ref();
            (
                attrs.and_then(|a| a.page_break_before) == Some(true),
                attrs.and_then(|a| a.page_break_before_run) == Some(true),
            )
        }
        _ => (false, false),
    }
}

/// Whether a paragraph must begin on a fresh page, however it asked.
pub fn paragraph_breaks_before(block: &LayoutBlock) -> bool {
    let (property, run) = paragraph_breaks_before_run(block);
    property || run
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
    fn counts_inherited_after_spacing_on_an_empty_member_like_placement_does() {
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

        assert_eq!(
            measure_keep_with_next_group(group.unwrap(), &measured),
            190.0
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
    fn does_not_advance_a_group_that_fits_with_inherited_empty_after_spacing() {
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
        assert_eq!(group_height, 190.0);

        // content height 864 (1056 - 2*96); the 620px filler leaves 244 available
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height,
            available_height: 244.0,
            page_content_height: 864.0,
            page_has_content: true,
        }));
    }

    fn make_heading_with_space_before(before: f64) -> LayoutBlock {
        paragraph(
            vec![text_run("Heading")],
            Some(ParagraphAttrs {
                keep_next: Some(true),
                keep_lines: Some(true),
                spacing: Some(ParagraphSpacing {
                    before: Some(before),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        )
    }

    #[test]
    fn witnesses_a_follower_widow_control_refuses_to_split_in_full() {
        let blocks = vec![
            make_heading_with_space_before(18.6667),
            make_paragraph_block("Follower", false),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(18.4), make_line(18.4)]),
            make_paragraph_measure(vec![
                make_line(16.8667),
                make_line(16.8667),
                make_line(16.8667),
            ]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan
            .groups_by_head
            .get(&0)
            .expect("group headed at block 0");

        // Three lines leave no legal internal break, so the witness is all of
        // them: 50.6001 + 18.6667 before + 36.8 heading.
        let group_height = measure_keep_with_next_group(group, &measured);
        assert!((group_height - 106.0668).abs() < 1e-3, "{group_height}");

        // oxi-en-legal-01 page 142: the pen stands at 885.68 px under a
        // 990 px content limit, so 104.32 px is left of an 835 px column.
        assert!(keep_with_next_group_must_advance(KeepWithNextFit {
            group_height,
            available_height: 104.32,
            page_content_height: 835.0,
            page_has_content: true,
        }));
    }

    #[test]
    fn witnesses_only_two_lines_of_a_follower_widow_control_can_split() {
        let blocks = vec![
            make_heading_with_space_before(18.6667),
            make_paragraph_block("Follower", false),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(18.4), make_line(18.4)]),
            make_paragraph_measure(vec![make_line(20.0); 4]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan
            .groups_by_head
            .get(&0)
            .expect("group headed at block 0");

        let group_height = measure_keep_with_next_group(group, &measured);
        assert!((group_height - 95.4667).abs() < 1e-3, "{group_height}");
    }

    #[test]
    fn witnesses_one_line_when_the_follower_turns_widow_control_off() {
        let blocks = vec![
            make_heading_with_space_before(18.6667),
            paragraph(
                vec![text_run("Follower")],
                Some(ParagraphAttrs {
                    widow_control: Some(false),
                    ..Default::default()
                }),
            ),
        ];
        let measures = vec![
            make_paragraph_measure(vec![make_line(18.4), make_line(18.4)]),
            make_paragraph_measure(vec![make_line(16.8667); 3]),
        ];
        let measured = to_measured_blocks(blocks, measures);

        let scan = analyze_keep_with_next(&measured);
        let group = scan
            .groups_by_head
            .get(&0)
            .expect("group headed at block 0");

        let group_height = measure_keep_with_next_group(group, &measured);
        assert!((group_height - 72.3334).abs() < 1e-3, "{group_height}");
    }
}
