//! The placement walk: measured blocks in, pages of positioned fragments out.
//!
//! One pass over the measured list, a per-kind placer for each block, the
//! paginator holding page and column state, and every look-ahead answered from
//! the prescan plan rather than by re-scanning. Contextual spacing
//! (`w:contextualSpacing`, ECMA-376 §17.3.1.9) is folded into adjacent
//! same-style paragraphs before the walk starts, recursing through table cells.
//!
//! Each block is processed in a fixed order: record a checkpoint if placement
//! stands at a pristine page start; force a page when the block carries
//! `w:pageBreakBefore` or opens with a hard `w:br w:type="page"` run, either
//! carrying the paragraph's space-before onto the new page; at the head of a
//! keep-with-next group, force a page when the group would otherwise straddle
//! the boundary; then dispatch to the placer.
//! A section break reads the *next* section's configuration and break type,
//! falling back to the current break's when the plan has no successor. Column
//! balancing runs over the range up to the next section break, both at the start
//! of the document and after each break that opens multi-column geometry, except
//! when that range ends with `nextColumn` and its successor needs the full band.
//!
//! Unknown block and run kinds raise `Unsupported` rather than being silently
//! dropped, and a measure whose kind disagrees with its block raises `Invalid`.
//!
//! Anchored images and floating text boxes overlay the page and never move the
//! pen; inline images, shapes, charts and inline text boxes consume their
//! measured bbox in flow. Paragraph placement is described on
//! `layout_paragraph`.
//!
//! Checkpoints are resume bookmarks at pristine page starts. They stay resident
//! and never appear in a serialized `Layout`. [`layout_document_incremental`]
//! restarts from the last checkpoint before the edit and stops as soon as
//! page-start geometry and the measured suffix converge with the retained
//! layout, so an edit late in a document does not re-place everything above it.

use crate::LayoutError;
use crate::hooks;
use crate::keep_together::{
    measure_keep_with_next_group_witnessing, paragraph_is_unbreakable, paragraph_widow_control,
};
use crate::page_flow::{PageFlowGeometry, Paginator};
use crate::paragraph_spacing::{
    apply_contextual_spacing_measured, get_spacing_after, get_spacing_before,
};
use crate::prescan::{LayoutPlan, SectionLayoutConfig, default_columns, prescan};
use crate::resolve_lines::{ResolvedLine, resolve_line_segments, utf16_len};
use crate::section_breaks::resolve_page_margins;
use crate::types::{
    BlockExtent, ChartBlock, ChartExtent, ChartFragment, Fragment, ImageBlock, ImageExtent,
    ImageFragment, ImageRunPosition, Input, Layout, LayoutBlock, MeasuredBlock, ParagraphBlock,
    ParagraphExtent, ParagraphFragment, Run, SectionBreakType, ShapeBlock, ShapeExtent,
    ShapeFragment, Size, TextBoxBlock, TextBoxExtent, TextBoxFragment,
};

/// Default page size (US Letter in pixels at 96 DPI).
const DEFAULT_PAGE_SIZE: Size = Size {
    w: 816.0,
    h: 1056.0,
};

/// A restart point at the beginning of a pristine page. These bookmarks stay
/// resident in Rust and are deliberately absent from serialized `Layout`.
#[derive(Debug, Clone, PartialEq)]
pub struct LayoutCheckpoint {
    pub block_index: usize,
    pub section_index: usize,
    pub page_index: usize,
    pub page_number: u32,
    pub flow: PageFlowGeometry,
}

/// Resident result of one placement pass, including execution metadata used
/// by the editor engine to propagate dirty-page damage.
#[derive(Debug)]
pub struct CheckpointedLayout {
    pub layout: Layout,
    pub checkpoints: Vec<LayoutCheckpoint>,
    pub placed_blocks: usize,
    pub rebuilt_page_start: usize,
    pub rebuilt_page_end: usize,
}

/// [`CheckpointedLayout`] of an incremental pass with the page ranges it placed afresh.
#[derive(Debug)]
pub struct IncrementalLayout {
    pub checkpointed: CheckpointedLayout,
    /// Ascending within `rebuilt_page_start..rebuilt_page_end`; the pages between
    /// them are retained.
    pub rebuilt_page_ranges: Vec<std::ops::Range<usize>>,
}

struct ConvergenceInput<'a, F: PartialEq> {
    previous_checkpoints: &'a [LayoutCheckpoint],
    previous_fingerprints: &'a [F],
    next_fingerprints: &'a [F],
    dirty_index: usize,
    /// Every dirty block, ascending, when the block count is unchanged.
    dirty: Option<&'a [usize]>,
    /// Whether placement may skip the clean pages between dirty blocks.
    skippable: bool,
    keep_with_next: &'a crate::keep_together::KeepWithNextScan,
    measured: &'a [MeasuredBlock],
    previous_pages: &'a [crate::types::Page],
    /// The last dirty block looked up, with the latest resumable checkpoint
    /// ahead of its restart.
    resume_before: std::cell::Cell<Option<(usize, Option<usize>)>>,
}

/// Where a placement walk met the retained layout.
enum Convergence {
    /// Every later block and page start matches: the retained suffix follows.
    Suffix {
        next: LayoutCheckpoint,
        previous: LayoutCheckpoint,
    },
    /// The retained pages up to `resume` follow unchanged, and placement
    /// resumes there ahead of the next dirty block.
    Skip {
        previous: LayoutCheckpoint,
        resume: LayoutCheckpoint,
        dirty_index: usize,
    },
}

fn pristine_paragraph_start(
    block_index: usize,
    measured: &[MeasuredBlock],
    page: &crate::types::Page,
) -> bool {
    let Some(block) = measured.get(block_index) else {
        return false;
    };
    let Some(Fragment::Paragraph(fragment)) = page.fragments.first() else {
        return false;
    };
    page.columns
        .as_ref()
        .is_none_or(|columns| columns.count <= 1.0)
        && page.float_bands.is_empty()
        && Some(&fragment.block_id) == block.block.block_id()
        && fragment.from_line == 0
        && fragment.carried_from_prev != Some(true)
        && (crate::break_policy::breaks_before_block(&block.block).is_some()
            || block_index
                .checked_sub(1)
                .and_then(|before| measured.get(before))
                .is_some_and(|before| {
                    matches!(
                        before.block,
                        LayoutBlock::PageBreak(_) | LayoutBlock::ColumnBreak(_)
                    )
                }))
}

impl<F: PartialEq> ConvergenceInput<'_, F> {
    fn retained_match(
        &self,
        checkpoint: &LayoutCheckpoint,
        opening_fragment_geometry: Option<&crate::page_flow::OpeningFragmentGeometry>,
    ) -> Option<Convergence> {
        if checkpoint.block_index <= self.dirty_index {
            return None;
        }
        let suffix_matches = match self.dirty {
            Some(dirty) => dirty
                .last()
                .is_none_or(|&last| last < checkpoint.block_index),
            None => {
                self.previous_fingerprints.get(checkpoint.block_index..)
                    == self.next_fingerprints.get(checkpoint.block_index..)
            }
        };
        if !suffix_matches && !self.skippable {
            return None;
        }
        // Checkpoints are in placement order, so those of one block are adjacent.
        let first = self
            .previous_checkpoints
            .partition_point(|previous| previous.block_index < checkpoint.block_index);
        let previous = self.previous_checkpoints[first..]
            .iter()
            .take_while(|previous| previous.block_index == checkpoint.block_index)
            .find(|previous| {
                previous.section_index == checkpoint.section_index
                    && previous.page_number == checkpoint.page_number
                    && self
                        .previous_pages
                        .get(previous.page_index)
                        .is_some_and(|page| {
                            page.opening_fragment_geometry.as_deref() == opening_fragment_geometry
                        })
                    && if suffix_matches {
                        flow_matches_from_page(
                            &previous.flow,
                            &checkpoint.flow,
                            checkpoint.page_number,
                        )
                    } else {
                        previous.flow == checkpoint.flow
                    }
            })?;
        if suffix_matches {
            return Some(Convergence::Suffix {
                next: checkpoint.clone(),
                previous: previous.clone(),
            });
        }
        if previous.page_index != checkpoint.page_index {
            return None;
        }
        let dirty = self.dirty?;
        let next_dirty =
            *dirty.get(dirty.partition_point(|&index| index < checkpoint.block_index))?;
        let resume_index = match self.resume_before.get() {
            Some((dirty_index, resume)) if dirty_index == next_dirty => resume,
            _ => {
                let restart = restart_index(self.keep_with_next, next_dirty);
                let end = self
                    .previous_checkpoints
                    .partition_point(|resume| resume.block_index < restart);
                let resume = self.previous_checkpoints[..end]
                    .iter()
                    .rposition(|resume| resumable(resume, self.measured, self.previous_pages));
                self.resume_before.set(Some((next_dirty, resume)));
                resume
            }
        };
        let resume = &self.previous_checkpoints[resume_index?];
        if resume.page_index <= previous.page_index {
            return None;
        }
        Some(Convergence::Skip {
            previous: previous.clone(),
            resume: resume.clone(),
            dirty_index: next_dirty,
        })
    }
}

fn flow_matches_from_page(
    previous: &PageFlowGeometry,
    next: &PageFlowGeometry,
    page_number: u32,
) -> bool {
    if previous == next {
        return true;
    }
    let mut flow = previous.clone();
    flow.footnote_reserved_heights
        .clone_from(&next.footnote_reserved_heights);
    if flow != *next {
        return false;
    }
    let previous = previous.footnote_reserved_heights.as_deref();
    let next = next.footnote_reserved_heights.as_deref();
    let reservation = |heights: Option<&std::collections::BTreeMap<String, f64>>, page: u32| {
        heights
            .and_then(|heights| heights.get(&page.to_string()).copied())
            .unwrap_or(0.0)
    };
    previous
        .into_iter()
        .flat_map(|heights| heights.keys())
        .chain(next.into_iter().flat_map(|heights| heights.keys()))
        .filter_map(|key| key.parse::<u32>().ok())
        .filter(|&page| page > 0 && page >= page_number)
        .all(|page| reservation(previous, page) == reservation(next, page))
}

/// Whether placement restarted at `checkpoint` lays out what follows as the walk
/// that recorded it did. A queued section geometry belongs to the page after the
/// checkpoint's; a block that opened a page whose geometry differs from the page
/// before was fitted under the earlier geometry; a table's rows were fitted
/// against the room left on the page before; and a floating or anchored object
/// that opened its page would be placed again from its anchor's offset. The first
/// page restarts from the document's origin instead.
fn resumable(
    checkpoint: &LayoutCheckpoint,
    measured: &[MeasuredBlock],
    pages: &[crate::types::Page],
) -> bool {
    let Some(MeasuredBlock {
        block: LayoutBlock::Paragraph(block),
        ..
    }) = measured.get(checkpoint.block_index)
    else {
        return false;
    };
    let Some(page) = pages.get(checkpoint.page_index) else {
        return false;
    };
    let mut fragments = page.fragments.iter().filter_map(|fragment| match fragment {
        Fragment::Paragraph(fragment) if fragment.block_id == block.id => Some(fragment),
        _ => None,
    });
    let starts_paragraph = fragments.next().is_some_and(|fragment| {
        fragment.from_line == 0 && fragment.carried_from_prev != Some(true)
    }) && (fragments.next().is_none()
        || pristine_paragraph_start(checkpoint.block_index, measured, page));
    let flow = &checkpoint.flow;
    checkpoint.page_index > 0
        && page.opening_fragment_geometry.is_none()
        && starts_paragraph
        && flow.pending_page_size.is_none()
        && flow.pending_margins.is_none()
        && flow.pending_columns.is_none()
}

/// Checkpoints in placement order: by block, then by page for a block that
/// opens two.
fn checkpoint_order(checkpoint: &LayoutCheckpoint) -> (usize, usize) {
    (checkpoint.block_index, checkpoint.page_index)
}

/// The block placement resumes before for a change at `dirty_index`. A dirty
/// block that opened a page may now start on the one before it (a removed page
/// break, a paragraph that now fits), and one inside or right after a
/// keep-with-next run can move that run's head, so resume strictly before either.
/// Right after a run counts even when the dirty block is no longer its
/// follower: a page break it gained releases the run from it.
fn restart_index(
    keep_with_next: &crate::keep_together::KeepWithNextScan,
    dirty_index: usize,
) -> usize {
    keep_with_next
        .groups_by_head
        .range(..=dirty_index)
        .rev()
        .take(2)
        .filter(|(_, group)| {
            group.members.contains(&dirty_index) || group.tail_index + 1 == dirty_index
        })
        .map(|(&head, _)| head)
        .min()
        .unwrap_or(dirty_index)
}

#[derive(Clone, Copy, PartialEq)]
enum Edge {
    Start,
    End,
}

fn run_boundary_pm_pos(run: Option<&Run>, char_offset: usize, edge: Edge) -> Option<f64> {
    let run = run?;

    if let Run::Text(r) = run {
        if let Some(pm_start) = r.pm_start {
            let clamped = char_offset.min(utf16_len(&r.text));
            return Some(pm_start + clamped as f64);
        }
        return if edge == Edge::End { r.pm_end } else { None };
    }

    if edge == Edge::End {
        if let Some(pm_end) = run.pm_end() {
            return Some(pm_end);
        }
        return run.pm_start().map(|pm| pm + 1.0);
    }
    run.pm_start()
}

/// Document range covered by lines `[from_line, to_line)`.
///
/// The paragraph's own bounds win at its true first and last line, so a
/// single-fragment paragraph keeps the range it declared; interior fragment
/// edges come from the head and tail runs of their boundary lines. Any edge
/// that stays unresolved falls back to the paragraph, and an end that would not
/// exceed its start is nudged forward by one so the range is never empty.
fn get_paragraph_fragment_pm_range(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    from_line: usize,
    to_line: usize,
) -> (Option<f64>, Option<f64>) {
    if measure.lines.is_empty() || from_line >= to_line {
        return (block.pm_start, block.pm_end);
    }

    let first_line = measure.lines.get(from_line);
    let last_line = measure.lines.get(to_line - 1);
    let first_run = first_line.and_then(|l| block.runs.get(l.head_run));
    let last_run = last_line.and_then(|l| block.runs.get(l.tail_run));
    let first_char = first_line.map_or(0, |l| l.head_char);
    let last_char = last_line.map_or(0, |l| l.tail_char);

    let mut pm_start = if from_line == 0 {
        block
            .pm_start
            .or_else(|| run_boundary_pm_pos(first_run, first_char, Edge::Start))
    } else {
        run_boundary_pm_pos(first_run, first_char, Edge::Start)
    };
    let mut pm_end = if to_line >= measure.lines.len() {
        block
            .pm_end
            .or_else(|| run_boundary_pm_pos(last_run, last_char, Edge::End))
    } else {
        run_boundary_pm_pos(last_run, last_char, Edge::End)
    };

    if pm_start.is_none() {
        pm_start = block.pm_start;
    }
    if pm_end.is_none() {
        pm_end = block.pm_end;
    }
    if let (Some(s), Some(e)) = (pm_start, pm_end)
        && e <= s
    {
        pm_end = Some(s + 1.0);
    }

    (pm_start, pm_end)
}

fn is_floating_wrap_type(wrap_type: Option<&str>) -> bool {
    matches!(
        wrap_type,
        Some("square") | Some("tight") | Some("through") | Some("behind") | Some("inFront")
    )
}

/// A text box floats when it declares float display, a floating wrap type, or
/// `topAndBottom` wrapping.
fn is_floating_text_box_block(block: &TextBoxBlock) -> bool {
    block.display_mode.as_deref() == Some("float")
        || is_floating_wrap_type(block.wrap_type.as_deref())
        || block.wrap_type.as_deref() == Some("topAndBottom")
}

/// The paginator a pass from the document's first block starts with.
fn origin_paginator(
    initial_config: &SectionLayoutConfig,
    plan: &LayoutPlan,
    options: &crate::types::LayoutOptions,
) -> Result<Paginator, LayoutError> {
    let mut paginator = Paginator::new(
        initial_config.page_size.clone(),
        initial_config.margins.clone(),
        initial_config
            .columns
            .clone()
            .unwrap_or_else(default_columns),
        options.footnote_reserved_heights.clone(),
    )?;
    paginator.set_section_page_margins(options.section_page_margins.clone().unwrap_or_default());
    paginator
        .set_section_page_float_bands(options.section_page_float_bands.clone().unwrap_or_default());
    if let Some(Some(restart)) = plan.section_page_restarts.first() {
        paginator.restart_page_numbering(restart.start);
    }
    Ok(paginator)
}

/// Converts measured blocks into positioned pages, discarding checkpoints.
pub fn layout_document(input: &mut Input) -> Result<Layout, LayoutError> {
    Ok(layout_document_checkpointed(input)?.layout)
}

/// Full placement while retaining clean page-start checkpoints in Rust.
pub fn layout_document_checkpointed(input: &mut Input) -> Result<CheckpointedLayout, LayoutError> {
    let measured = &mut input.measured;
    let options = &input.options;

    let page_size = options.page_size.clone().unwrap_or(DEFAULT_PAGE_SIZE);
    let margins = resolve_page_margins(options.margins.as_ref());
    let final_page_size = options
        .final_page_size
        .clone()
        .unwrap_or_else(|| page_size.clone());
    let final_margins = options
        .final_margins
        .clone()
        .unwrap_or_else(|| margins.clone());

    let content_width = page_size.w - margins.left - margins.right;
    if content_width <= 0.0 {
        return Err(LayoutError::Invalid(
            "layoutDocument: page size and margins yield no content area".into(),
        ));
    }

    let body_config = SectionLayoutConfig {
        page_size: page_size.clone(),
        margins: margins.clone(),
        columns: options.columns.clone(),
    };
    let final_config = SectionLayoutConfig {
        page_size: final_page_size,
        margins: final_margins,
        columns: options.columns.clone(),
    };

    // mutate spacing attrs before anything reads them — the keep-with-next
    // group height must see contextual-spacing suppression (§17.3.1.9)
    apply_contextual_spacing_measured(measured);

    let mut plan = prescan(
        measured,
        &body_config,
        final_config,
        options.body_break_type,
    )?;
    plan.section_page_restarts = options.section_page_restarts.clone().unwrap_or_default();

    let initial_config = plan.section_configs.first().cloned().unwrap_or(body_config);

    let mut paginator = origin_paginator(&initial_config, &plan, options)?;

    let placement = place::<u64>(
        measured,
        &plan,
        &mut paginator,
        &initial_config,
        0,
        0,
        0,
        None,
    )?;

    // an empty document still yields page 1
    if paginator.pages.is_empty() {
        paginator.get_current();
    }

    let page_count = paginator.pages.len();
    Ok(CheckpointedLayout {
        layout: Layout {
            page_size,
            pages: paginator.pages,
            columns: options.columns.clone(),
            headers: None,
            footers: None,
            page_gap: options.page_gap,
            partial: false,
            cached_page_totals: false,
        },
        checkpoints: placement.checkpoints,
        placed_blocks: placement.placed_blocks,
        rebuilt_page_start: 0,
        rebuilt_page_end: page_count,
    })
}

/// Resume placement from the last clean checkpoint preceding `dirty_index`,
/// then stop as soon as page-start geometry and the measured suffix converge
/// with the retained layout. Callers must conservatively gate unsupported
/// dependency shapes (floats, notes, structural edits) before entering here.
pub fn layout_document_incremental<F: PartialEq>(
    input: &mut Input,
    previous_layout: &mut Layout,
    previous_checkpoints: &[LayoutCheckpoint],
    previous_fingerprints: &[F],
    next_fingerprints: &[F],
    dirty_index: usize,
) -> Result<CheckpointedLayout, LayoutError> {
    layout_document_incremental_ranges(
        input,
        previous_layout,
        previous_checkpoints,
        previous_fingerprints,
        next_fingerprints,
        dirty_index,
    )
    .map(|run| run.checkpointed)
}

/// [`layout_document_incremental`], reporting the page ranges it placed afresh.
pub fn layout_document_incremental_ranges<F: PartialEq>(
    input: &mut Input,
    previous_layout: &mut Layout,
    previous_checkpoints: &[LayoutCheckpoint],
    previous_fingerprints: &[F],
    next_fingerprints: &[F],
    dirty_index: usize,
) -> Result<IncrementalLayout, LayoutError> {
    let options = &input.options;
    let page_size = options.page_size.clone().unwrap_or(DEFAULT_PAGE_SIZE);
    let margins = resolve_page_margins(options.margins.as_ref());
    let final_page_size = options
        .final_page_size
        .clone()
        .unwrap_or_else(|| page_size.clone());
    let final_margins = options
        .final_margins
        .clone()
        .unwrap_or_else(|| margins.clone());
    if page_size.w - margins.left - margins.right <= 0.0 {
        return Err(LayoutError::Invalid(
            "layoutDocument: page size and margins yield no content area".into(),
        ));
    }

    let body_config = SectionLayoutConfig {
        page_size: page_size.clone(),
        margins,
        columns: options.columns.clone(),
    };
    let final_config = SectionLayoutConfig {
        page_size: final_page_size,
        margins: final_margins,
        columns: options.columns.clone(),
    };
    apply_contextual_spacing_measured(&mut input.measured);
    let mut plan = prescan(
        &input.measured,
        &body_config,
        final_config,
        options.body_break_type,
    )?;
    plan.section_page_restarts = options.section_page_restarts.clone().unwrap_or_default();
    let initial_config = plan.section_configs.first().cloned().unwrap_or(body_config);
    let restart = restart_index(&plan.keep_with_next, dirty_index);
    // Balancing weighs all of a section's content, so an edit anywhere in a
    // section whose columns balance resumes before the section opens them.
    let dirty_section = plan
        .break_indices
        .partition_point(|&index| index < dirty_index);
    let restart = if section_balances_columns(&plan, dirty_section, &initial_config) {
        restart.min(section_start(&plan, dirty_section))
    } else {
        restart
    };
    // Without a resumable checkpoint ahead of the change, placement starts
    // afresh at the document's origin, still converging with the retained layout.
    let resume = previous_checkpoints.iter().rev().find(|checkpoint| {
        checkpoint.block_index < restart
            && resumable(checkpoint, &input.measured, &previous_layout.pages)
    });
    let resume_page = resume.map_or(0, |resume| resume.page_index);
    let dirty: Option<Vec<usize>> =
        (previous_fingerprints.len() == next_fingerprints.len()).then(|| {
            previous_fingerprints
                .iter()
                .zip(next_fingerprints)
                .enumerate()
                .filter(|(_, (previous, next))| previous != next)
                .map(|(index, _)| index)
                .collect()
        });
    // Clean stretches between dirty blocks are skipped only where a page start
    // alone decides what follows: one column throughout, no reserved note space,
    // and no later section break that changed.
    let skippable = dirty.as_ref().is_some_and(|dirty| !dirty.is_empty())
        && plan
            .section_configs
            .iter()
            .chain([&initial_config])
            .all(|config| {
                config
                    .columns
                    .as_ref()
                    .is_none_or(|columns| columns.count <= 1.0)
            })
        && options
            .footnote_reserved_heights
            .as_ref()
            .is_none_or(|heights| heights.is_empty())
        && !dirty
            .iter()
            .flatten()
            .skip(1)
            .any(|&index| matches!(input.measured[index].block, LayoutBlock::SectionBreak(_)));

    // Place each dirty stretch in turn without touching the retained pages, so
    // a failure leaves the caller's layout as it was.
    let mut segments: Vec<(usize, Vec<crate::types::Page>, Option<Convergence>)> = Vec::new();
    let mut checkpoints: Vec<_> = resume.map_or_else(Vec::new, |resume| {
        previous_checkpoints
            .iter()
            .filter(|checkpoint| checkpoint_order(checkpoint) < checkpoint_order(resume))
            .cloned()
            .collect()
    });
    let mut placed_blocks = 0;
    let mut start = resume.cloned();
    let mut segment_dirty = dirty_index;
    loop {
        let (block_index, section_index, page_index, mut paginator) = match &start {
            Some(start) => {
                let mut paginator = Paginator::resume_in_section(
                    &start.flow,
                    start.page_number,
                    start.section_index,
                    options.footnote_reserved_heights.clone(),
                )?;
                paginator.set_section_page_margins(
                    options.section_page_margins.clone().unwrap_or_default(),
                );
                paginator.set_section_page_float_bands(
                    options.section_page_float_bands.clone().unwrap_or_default(),
                );
                (
                    start.block_index,
                    start.section_index,
                    start.page_index,
                    paginator,
                )
            }
            None => (0, 0, 0, origin_paginator(&initial_config, &plan, options)?),
        };
        if segments.is_empty() && !checkpoints.is_empty() {
            let reservations = paginator.snapshot_geometry().footnote_reserved_heights;
            for checkpoint in &mut checkpoints {
                checkpoint
                    .flow
                    .footnote_reserved_heights
                    .clone_from(&reservations);
            }
        }
        let convergence = ConvergenceInput {
            previous_checkpoints,
            previous_fingerprints,
            next_fingerprints,
            dirty_index: segment_dirty,
            dirty: dirty.as_deref(),
            skippable,
            keep_with_next: &plan.keep_with_next,
            measured: &input.measured,
            previous_pages: &previous_layout.pages,
            resume_before: std::cell::Cell::new(None),
        };
        let placement = place(
            &input.measured,
            &plan,
            &mut paginator,
            &initial_config,
            block_index,
            section_index,
            page_index,
            Some(&convergence),
        )?;
        placed_blocks += placement.placed_blocks;
        checkpoints.extend(placement.checkpoints);
        let next = match &placement.converged {
            Some(Convergence::Skip {
                previous,
                resume,
                dirty_index,
            }) => {
                let bound = |checkpoint: &LayoutCheckpoint| {
                    previous_checkpoints.partition_point(|retained| {
                        checkpoint_order(retained) < checkpoint_order(checkpoint)
                    })
                };
                checkpoints
                    .extend_from_slice(&previous_checkpoints[bound(previous)..bound(resume)]);
                segment_dirty = *dirty_index;
                Some(resume.clone())
            }
            _ => None,
        };
        segments.push((page_index, paginator.pages, placement.converged));
        match next {
            Some(resume) => start = Some(resume),
            None => break,
        }
    }

    let mut retained = std::mem::take(&mut previous_layout.pages)
        .into_iter()
        .enumerate()
        .peekable();
    let mut take_retained = |from: usize, to: usize, pages: &mut Vec<crate::types::Page>| {
        while let Some((index, _)) = retained.peek() {
            if *index >= to {
                break;
            }
            let (index, page) = retained.next().expect("peeked page");
            if index >= from {
                pages.push(page);
            }
        }
    };
    let mut pages = Vec::new();
    take_retained(0, resume_page, &mut pages);
    let mut rebuilt_page_ranges = Vec::with_capacity(segments.len());
    let mut reused_ranges = Vec::new();
    for (start_page, mut placed, converged) in segments {
        debug_assert_eq!(pages.len(), start_page);
        let rebuilt_start = pages.len();
        pages.append(&mut placed);
        rebuilt_page_ranges.push(rebuilt_start..pages.len());
        let reused_start = pages.len();
        match converged {
            Some(Convergence::Skip {
                previous, resume, ..
            }) => take_retained(previous.page_index, resume.page_index, &mut pages),
            Some(Convergence::Suffix { next, previous }) => {
                debug_assert_eq!(pages.len(), next.page_index);
                take_retained(previous.page_index, usize::MAX, &mut pages);
                let page_shift = next.page_index as isize - previous.page_index as isize;
                checkpoints.extend(
                    previous_checkpoints
                        .iter()
                        .filter(|checkpoint| {
                            checkpoint_order(checkpoint) >= checkpoint_order(&previous)
                        })
                        .cloned()
                        .map(|mut checkpoint| {
                            checkpoint.page_index =
                                (checkpoint.page_index as isize + page_shift) as usize;
                            checkpoint
                                .flow
                                .footnote_reserved_heights
                                .clone_from(&next.flow.footnote_reserved_heights);
                            checkpoint
                        }),
                );
            }
            None => {}
        }
        reused_ranges.push(reused_start..pages.len());
    }
    refresh_reused_page_ranges(&mut pages, &reused_ranges, &input.measured);

    Ok(IncrementalLayout {
        checkpointed: CheckpointedLayout {
            layout: Layout {
                page_size,
                pages,
                columns: options.columns.clone(),
                headers: None,
                footers: None,
                page_gap: options.page_gap,
                partial: false,
                cached_page_totals: false,
            },
            checkpoints,
            placed_blocks,
            rebuilt_page_start: resume_page,
            rebuilt_page_end: rebuilt_page_ranges
                .last()
                .map_or(resume_page, |range| range.end),
        },
        rebuilt_page_ranges,
    })
}

struct PlacementOutcome {
    checkpoints: Vec<LayoutCheckpoint>,
    placed_blocks: usize,
    converged: Option<Convergence>,
}

fn break_type_after_section(plan: &LayoutPlan, section_index: usize) -> Option<SectionBreakType> {
    plan.section_break_types
        .get(section_index + 1)
        .copied()
        .flatten()
        .or_else(|| {
            plan.section_break_types
                .get(section_index)
                .copied()
                .flatten()
        })
}

fn section_ends_with_next_column(plan: &LayoutPlan, section_index: usize) -> bool {
    plan.break_indices.get(section_index).is_some()
        && break_type_after_section(plan, section_index) == Some(SectionBreakType::NextColumn)
}

/// Whether placement balances a section's columns where its column region opens.
fn section_balances_columns(
    plan: &LayoutPlan,
    section_index: usize,
    initial_config: &SectionLayoutConfig,
) -> bool {
    plan.section_configs
        .get(section_index)
        .unwrap_or(initial_config)
        .columns
        .as_ref()
        .map_or(1.0, |columns| columns.count)
        > 1.0
        && !section_ends_with_next_column(plan, section_index)
}

/// The index of a section's first block.
fn section_start(plan: &LayoutPlan, section_index: usize) -> usize {
    section_index
        .checked_sub(1)
        .and_then(|previous| plan.break_indices.get(previous))
        .map_or(0, |section_break| section_break + 1)
}

/// The block walk itself, per the module's ordering rules. Returns early once
/// a checkpoint matches the retained layout, which is how incremental placement
/// detects convergence.
fn place<F: PartialEq>(
    measured: &[MeasuredBlock],
    plan: &LayoutPlan,
    paginator: &mut Paginator,
    initial_config: &SectionLayoutConfig,
    start_index: usize,
    mut section_idx: usize,
    page_index_offset: usize,
    convergence: Option<&ConvergenceInput<'_, F>>,
) -> Result<PlacementOutcome, LayoutError> {
    let mut checkpoints = Vec::new();
    let mut placed_blocks = 0usize;

    // Balancing belongs to the page a section's column region opens on: a
    // pass from the document start balances the first section, and a resumed
    // pass rebalances only when its checkpoint's page is that page.
    let balances = if start_index == 0 {
        section_balances_columns(plan, 0, initial_config)
    } else {
        paginator.balances_region()
    };
    if balances {
        hooks::balance_terminal_continuous_text_columns(
            measured,
            paginator,
            section_start(plan, section_idx),
            plan.break_indices
                .get(section_idx)
                .copied()
                .unwrap_or(measured.len()),
        )?;
        paginator.mark_balanced_region();
    }

    for (i, mb) in measured.iter().enumerate().skip(start_index) {
        let mut checkpointed_page = None;
        if let Some((page_index, page_number, flow)) = paginator.clean_page_start() {
            let checkpoint = LayoutCheckpoint {
                block_index: i,
                section_index: section_idx,
                page_index: page_index_offset + page_index,
                page_number,
                flow,
            };
            if let Some(converged) =
                convergence.and_then(|value| value.retained_match(&checkpoint, None))
            {
                paginator.pages.truncate(page_index);
                return Ok(PlacementOutcome {
                    checkpoints,
                    placed_blocks,
                    converged: Some(converged),
                });
            }
            checkpointed_page = Some(page_index);
            checkpoints.push(checkpoint);
        }
        let fragments_before = paginator.page_fragment_counts();
        // pageBreakBefore, or a hard page-break run, forces a fresh page and
        // keeps the paragraph's space-before net of the previous space-after
        if let Some(authored) = hooks::breaks_before_block(&mb.block)? {
            paginator.force_authored_page_break(authored.keeps_leading_spacing());
        }

        // at the head of a keep-with-next group, move to a fresh column when the
        // whole group would otherwise straddle the boundary
        if let Some(group) = plan.keep_with_next.groups_by_head.get(&i)
            && !plan.keep_with_next.interior_members.contains(&i)
        {
            let state_idx = paginator.get_current();
            let page_content_height =
                paginator.state(state_idx).content_limit - paginator.state(state_idx).content_top;
            let page_has_content = paginator.page_fragment_count(state_idx) > 0;
            // between float bands a table row's first slice may not share a gap
            // with the run, so a table follower keeps its whole first row there
            let split_first_row = !paginator.has_float_bands();
            let group_height = measure_keep_with_next_group_witnessing(
                group,
                measured,
                |before| paginator.leading_spacing(before),
                paginator.state(state_idx).deferred_spacing,
                page_content_height,
                split_first_row,
            );
            let fresh_page_height = measure_keep_with_next_group_witnessing(
                group,
                measured,
                |_| 0.0,
                0.0,
                page_content_height,
                split_first_row,
            );
            let must_advance = hooks::keep_with_next_group_must_advance_from(
                group_height,
                fresh_page_height,
                paginator.get_available_height(),
                page_content_height,
                page_has_content,
            )?;
            if must_advance {
                if paginator.has_float_bands() {
                    paginator.ensure_fits(group_height);
                } else {
                    // advance until a column holds the run or a fresh page opens
                    loop {
                        let idx = paginator.advance_for_overflow();
                        if paginator.state(idx).column_index == 0
                            || fresh_page_height <= paginator.get_available_height()
                        {
                            break;
                        }
                    }
                }
            }
        }

        match &mb.block {
            LayoutBlock::Paragraph(block) => {
                let BlockExtent::Paragraph(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutParagraph: expected paragraph measure".into(),
                    ));
                };
                layout_paragraph(block, measure, paginator)?;
            }

            LayoutBlock::Table(block) => {
                let BlockExtent::Table(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutTable: expected table measure".into(),
                    ));
                };
                if block.floating.is_some() {
                    let content_width = paginator.get_content_width();
                    hooks::layout_floating_table(block, measure, paginator, content_width)?;
                } else {
                    hooks::layout_table(block, measure, paginator)?;
                }
            }

            LayoutBlock::Image(block) => {
                let BlockExtent::Image(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutImage: expected image measure".into(),
                    ));
                };
                layout_image(block, measure, paginator);
            }

            LayoutBlock::Shape(block) => {
                let BlockExtent::Shape(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutShape: expected shape measure".into(),
                    ));
                };
                layout_shape(block, measure, paginator);
            }

            LayoutBlock::Chart(block) => {
                let BlockExtent::Chart(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutChart: expected chart measure".into(),
                    ));
                };
                layout_chart(block, measure, paginator);
            }

            LayoutBlock::TextBox(block) => {
                let BlockExtent::TextBox(measure) = &mb.measure else {
                    return Err(LayoutError::Invalid(
                        "layoutTextBox: expected textBox measure".into(),
                    ));
                };
                layout_text_box(block, measure, paginator);
            }

            LayoutBlock::PageBreak(_) => {
                paginator.force_authored_page_break(false);
            }

            LayoutBlock::ColumnBreak(_) => {
                paginator.force_column_break();
            }

            LayoutBlock::SectionBreak(block) => {
                // use the NEXT section's columns; for break type, prefer the
                // next section's but fall back to the current break's
                let next_type = break_type_after_section(plan, section_idx);
                let restart = plan
                    .section_page_restarts
                    .get(section_idx + 1)
                    .copied()
                    .flatten();
                let next_section_config = plan
                    .section_configs
                    .get(section_idx + 1)
                    .cloned()
                    .unwrap_or_else(|| initial_config.clone());
                let restart_starts_page = restart.is_some()
                    && crate::section_breaks::restart_starts_page(
                        paginator,
                        &next_section_config,
                        next_type,
                    );
                let next_type = match (next_type, restart) {
                    (Some(SectionBreakType::OddPage | SectionBreakType::EvenPage), _) => next_type,
                    (_, Some(restart)) if restart.align_parity && restart_starts_page => {
                        Some(if paginator.physical_parity_is_odd(restart.start) {
                            SectionBreakType::OddPage
                        } else {
                            SectionBreakType::EvenPage
                        })
                    }
                    _ => next_type,
                };
                let opened_column_region =
                    hooks::handle_section_break(block, paginator, &next_section_config, next_type)?;
                paginator.set_section_index(section_idx + 1);
                if let Some(restart) = restart
                    && restart_starts_page
                {
                    paginator.restart_page_numbering(restart.start);
                }

                let next_break_index = plan.break_indices.get(section_idx + 1).copied();
                if opened_column_region
                    && next_section_config
                        .columns
                        .as_ref()
                        .map_or(1.0, |c| c.count)
                        > 1.0
                    && !section_ends_with_next_column(plan, section_idx + 1)
                {
                    hooks::balance_terminal_continuous_text_columns(
                        measured,
                        paginator,
                        i + 1,
                        next_break_index.unwrap_or(measured.len()),
                    )?;
                    paginator.mark_balanced_region();
                }

                section_idx += 1;
            }

            LayoutBlock::Unsupported => {
                return Err(LayoutError::Unsupported("unknown block kind".into()));
            }
        }

        placed_blocks += 1;
        let fragments_after = paginator.page_fragment_counts();
        let first_changed_page = fragments_after
            .iter()
            .enumerate()
            .find_map(|(page, count)| {
                let before = fragments_before.get(page).copied().unwrap_or(0);
                (*count > before).then_some((page, before))
            });
        if let Some((page_index, 0)) = first_changed_page
            && checkpointed_page != Some(page_index)
            && (!matches!(mb.block, LayoutBlock::Paragraph(_))
                || fragments_after[page_index] == 1
                || pristine_paragraph_start(i, measured, &paginator.pages[page_index]))
            && paginator
                .current_page_start()
                .is_some_and(|(current, _, _)| current == page_index)
        {
            let (_, page_number, flow) = paginator
                .current_page_start()
                .expect("current page was checked above");
            let checkpoint = LayoutCheckpoint {
                block_index: i,
                section_index: section_idx,
                page_index: page_index_offset + page_index,
                page_number,
                flow,
            };
            if let Some(converged) = convergence.and_then(|value| {
                value.retained_match(
                    &checkpoint,
                    paginator.pages[page_index]
                        .opening_fragment_geometry
                        .as_deref(),
                )
            }) {
                paginator.pages.truncate(page_index);
                return Ok(PlacementOutcome {
                    checkpoints,
                    placed_blocks,
                    converged: Some(converged),
                });
            }
            checkpoints.push(checkpoint);
        }
    }

    Ok(PlacementOutcome {
        checkpoints,
        placed_blocks,
        converged: None,
    })
}

fn block_id_key(id: &crate::types::BlockId) -> String {
    serde_json::to_string(id).expect("block ids always serialize")
}

/// Retained suffix pages keep their geometry but absolute document positions move
/// after an earlier edit. Refresh fragment ranges and resolved run slices from
/// the new measured arena before the display list consumes them.
fn refresh_reused_page_ranges(
    pages: &mut [crate::types::Page],
    ranges: &[std::ops::Range<usize>],
    measured: &[MeasuredBlock],
) {
    if ranges.iter().all(std::ops::Range::is_empty) {
        return;
    }
    let blocks: std::collections::HashMap<_, _> = measured
        .iter()
        .filter_map(|measured| {
            measured
                .block
                .block_id()
                .map(|id| (block_id_key(id), measured))
        })
        .collect();
    for range in ranges {
        refresh_reused_pages(&mut pages[range.clone()], &blocks);
    }
}

fn refresh_reused_pages(
    pages: &mut [crate::types::Page],
    blocks: &std::collections::HashMap<String, &MeasuredBlock>,
) {
    for page in pages {
        for fragment in &mut page.fragments {
            let key = match fragment {
                Fragment::Paragraph(fragment) => block_id_key(&fragment.block_id),
                Fragment::Table(fragment) => block_id_key(&fragment.block_id),
                Fragment::Image(fragment) => block_id_key(&fragment.block_id),
                Fragment::Shape(fragment) => block_id_key(&fragment.block_id),
                Fragment::Chart(fragment) => block_id_key(&fragment.block_id),
                Fragment::TextBox(fragment) => block_id_key(&fragment.block_id),
            };
            let Some(measured) = blocks.get(&key) else {
                continue;
            };
            match (fragment, &measured.block, &measured.measure) {
                (
                    Fragment::Paragraph(fragment),
                    LayoutBlock::Paragraph(block),
                    BlockExtent::Paragraph(extent),
                ) => {
                    (fragment.pm_start, fragment.pm_end) = get_paragraph_fragment_pm_range(
                        block,
                        extent,
                        fragment.from_line,
                        fragment.to_line,
                    );
                    fragment.resolved_lines = Some(build_resolved_lines(
                        block,
                        extent,
                        fragment.from_line,
                        fragment.to_line,
                    ));
                }
                (Fragment::Table(fragment), LayoutBlock::Table(block), _) => {
                    fragment.pm_start = block.pm_start;
                    fragment.pm_end = block.pm_end;
                }
                (Fragment::Image(fragment), LayoutBlock::Image(block), _) => {
                    fragment.pm_start = block.pm_start;
                    fragment.pm_end = block.pm_end;
                }
                (Fragment::Shape(fragment), LayoutBlock::Shape(block), _) => {
                    fragment.pm_start = block.pm_start;
                    fragment.pm_end = block.pm_end;
                    fragment.doc_start = block.doc_start;
                    fragment.doc_end = block.doc_end;
                }
                (Fragment::Chart(fragment), LayoutBlock::Chart(block), _) => {
                    fragment.pm_start = block.pm_start;
                    fragment.pm_end = block.pm_end;
                    fragment.doc_start = block.doc_start;
                    fragment.doc_end = block.doc_end;
                }
                (Fragment::TextBox(fragment), LayoutBlock::TextBox(block), _) => {
                    fragment.pm_start = block.pm_start;
                    fragment.pm_end = block.pm_end;
                }
                _ => {}
            }
        }
    }
}

// ---------------------------------------------------------------------------
// per-kind placers
// ---------------------------------------------------------------------------

/// Materializes resolved run segments for a fragment line range.
fn build_resolved_lines(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    from_line: usize,
    to_line: usize,
) -> Vec<ResolvedLine> {
    let mut resolved = Vec::new();
    for line_index in from_line..to_line {
        let Some(line) = measure.lines.get(line_index) else {
            continue;
        };
        resolved.push(ResolvedLine {
            segments: resolve_line_segments(&block.runs, line),
        });
    }
    resolved
}

fn paragraph_fragment_height(before: f64, lines_height: f64) -> f64 {
    before + lines_height
}

/// Places a paragraph's measured lines, splitting into carried fragments
/// whenever the page or column runs out of room.
///
/// Lines are fitted greedily, and a fragment always takes at least one line so
/// an oversized line cannot stall the walk. A line's `floatSkipBefore` counts
/// toward the fragment height, which is what makes following blocks flow below
/// the float instead of over it. A paragraph with no measured lines still emits
/// a zero-height fragment, because its spacing must still advance the pen.
///
/// Two rules can move lines before they are placed. `w:keepLines` advances to a
/// fresh column when the whole paragraph fits a column but not the space left
/// here. Widow and orphan control keeps two- and three-line paragraphs together
/// unless they turn `w:widowControl` off. For longer paragraphs, a lone opening
/// line moves the paragraph on when two lines would fit there, and a lone trailing line is
/// avoided by pushing one more line down, provided the fragment keeps more than
/// two.
///
/// Spacing before is charged to the first fragment only and spacing after to
/// the last, and each fragment carries its own document range and resolved run
/// slices.
fn layout_paragraph(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    paginator: &mut Paginator,
) -> Result<(), LayoutError> {
    // an unknown run kind can't be re-emitted faithfully in resolved lines
    if block.runs.iter().any(|r| matches!(r, Run::Unsupported)) {
        return Err(LayoutError::Unsupported("unknown run kind".into()));
    }

    let lines = &measure.lines;
    if lines.is_empty() {
        // no measured lines: a zero-height fragment still advances the pen by
        // its spacing
        let space_before = get_spacing_before(block);
        let space_after = get_spacing_after(block);
        let state_idx = paginator.get_current();
        let column_index = paginator.state(state_idx).column_index;
        let pen_y = paginator.state(state_idx).pen_y;

        let fragment = Fragment::Paragraph(ParagraphFragment {
            block_id: block.id.clone(),
            x: paginator.get_column_x(column_index),
            y: pen_y + paginator.leading_spacing(space_before),
            width: paginator.get_content_width(),
            height: 0.0,
            from_line: 0,
            to_line: 0,
            pm_start: block.pm_start,
            pm_end: block.pm_end,
            carried_from_prev: None,
            carried_to_next: None,
            resolved_lines: Some(Vec::new()),
        });

        paginator.add_fragment(fragment, 0.0, space_before, space_after);
        return Ok(());
    }

    let space_before = get_spacing_before(block);
    let space_after = get_spacing_after(block);
    let paragraph_height = lines.iter().fold(0.0, |sum, line| {
        sum + line.line_height + line.float_skip_before.unwrap_or(0.0)
    });
    let widow_control = paragraph_widow_control(block, measure);

    if paragraph_is_unbreakable(block, measure) {
        let state_idx = paginator.get_current();
        let state = paginator.state(state_idx);
        let required = paginator
            .leading_spacing(space_before)
            .max(state.deferred_spacing)
            + paragraph_height;
        let capacity = paginator.get_column_capacity();
        if paragraph_height <= capacity && required > paginator.get_available_height() {
            paginator.ensure_fits(required);
        }
    }

    let mut current_line_index = 0usize;

    while current_line_index < lines.len() {
        if paginator.has_float_bands() {
            let state_idx = paginator.get_current();
            let before = if current_line_index == 0 {
                paginator
                    .leading_spacing(space_before)
                    .max(paginator.state(state_idx).deferred_spacing)
            } else {
                0.0
            };
            paginator.ensure_fits(paragraph_fragment_height(
                before,
                lines[current_line_index].line_height
                    + lines[current_line_index].float_skip_before.unwrap_or(0.0),
            ));
        }
        let state_idx = paginator.get_current();
        let deferred_spacing = paginator.state(state_idx).deferred_spacing;
        let column_index = paginator.state(state_idx).column_index;

        // Reserve leading space before fitting the first fragment.
        let reserved_before = if current_line_index == 0 {
            paginator
                .leading_spacing(space_before)
                .max(deferred_spacing)
        } else {
            0.0
        };
        let available_height = paginator.get_available_height();
        let has_float_bands = paginator.has_float_bands();
        let (fit_before, available_for_lines) = if has_float_bands {
            (reserved_before, available_height)
        } else {
            (0.0, available_height - reserved_before)
        };

        // greedy fit; a fragment always takes at least one line
        let mut lines_height = 0.0f64;
        let mut fitting_lines = 0usize;

        for line in &lines[current_line_index..] {
            // floatSkipBefore counts toward fragment height so following
            // blocks flow below the float, not over it
            let line_height = line.line_height + line.float_skip_before.unwrap_or(0.0);
            let total_with_line = lines_height + line_height;

            if paragraph_fragment_height(fit_before, total_with_line) <= available_for_lines
                || fitting_lines == 0
            {
                lines_height = total_with_line;
                fitting_lines += 1;
            } else {
                break;
            }
        }

        let remaining_after = lines.len() - (current_line_index + fitting_lines);
        let mut pushed_widow = false;
        if widow_control && remaining_after > 0 {
            if current_line_index == 0 && fitting_lines == 1 {
                let capacity = paginator.get_column_capacity();
                let first_two_height = lines.iter().take(2).fold(0.0, |sum, line| {
                    if has_float_bands {
                        sum + (line.line_height + line.float_skip_before.unwrap_or(0.0))
                    } else {
                        sum + line.line_height + line.float_skip_before.unwrap_or(0.0)
                    }
                });
                let required = paragraph_fragment_height(reserved_before, first_two_height);
                if required <= capacity {
                    let pen_y = paginator.state(state_idx).pen_y;
                    let next_idx = if has_float_bands {
                        paginator.ensure_fits(required)
                    } else {
                        paginator.advance_for_overflow()
                    };
                    let next = paginator.state(next_idx);
                    if next_idx != state_idx
                        || next.column_index != column_index
                        || next.pen_y > pen_y
                    {
                        continue;
                    }
                }
            }
            if remaining_after == 1 && fitting_lines > 2 {
                fitting_lines -= 1;
                // at a float band the space below it still takes the line, and
                // balancing chose its column depth with the line kept here
                let tail: f64 = lines[current_line_index + fitting_lines..]
                    .iter()
                    .map(|line| line.line_height + line.float_skip_before.unwrap_or(0.0))
                    .sum();
                pushed_widow = !has_float_bands
                    && !paginator.balances_region()
                    && tail <= paginator.get_column_capacity();
                let removed = &lines[current_line_index + fitting_lines];
                lines_height -= removed.line_height + removed.float_skip_before.unwrap_or(0.0);
            }
        }

        let is_first_fragment = current_line_index == 0;
        let is_last_fragment = current_line_index + fitting_lines >= lines.len();
        let effective_space_before = if is_first_fragment { space_before } else { 0.0 };
        let effective_space_after = if is_last_fragment { space_after } else { 0.0 };
        let (pm_start, pm_end) = get_paragraph_fragment_pm_range(
            block,
            measure,
            current_line_index,
            current_line_index + fitting_lines,
        );

        let fragment = Fragment::Paragraph(ParagraphFragment {
            block_id: block.id.clone(),
            x: paginator.get_column_x(column_index),
            y: 0.0, // set by add_fragment
            width: paginator.get_content_width(),
            height: lines_height,
            from_line: current_line_index,
            to_line: current_line_index + fitting_lines,
            pm_start,
            pm_end,
            carried_from_prev: Some(!is_first_fragment),
            carried_to_next: Some(!is_last_fragment),
            resolved_lines: Some(build_resolved_lines(
                block,
                measure,
                current_line_index,
                current_line_index + fitting_lines,
            )),
        });

        paginator.add_fragment(
            fragment,
            lines_height,
            effective_space_before,
            effective_space_after,
        );

        current_line_index += fitting_lines;

        // leftover lines: move the pen to a column/page with room for the next;
        // a line widow control pushed down still fits here, so break anyway
        if pushed_widow {
            paginator.advance_for_overflow();
        } else if current_line_index < lines.len() {
            paginator.ensure_fits(lines[current_line_index].line_height);
        }
    }

    Ok(())
}

/// Places inline images in flow and anchored images over the page.
fn layout_image(block: &ImageBlock, measure: &ImageExtent, paginator: &mut Paginator) {
    if block
        .anchor
        .as_ref()
        .and_then(|a| a.is_anchored)
        .unwrap_or(false)
    {
        layout_anchored_image(block, measure, paginator);
        return;
    }

    let state_idx = paginator.ensure_fits(measure.height);
    let column_index = paginator.state(state_idx).column_index;

    let fragment = Fragment::Image(ImageFragment {
        block_id: block.id.clone(),
        x: paginator.get_column_x(column_index),
        y: 0.0, // set by add_fragment
        width: measure.width,
        height: measure.height,
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        is_anchored: None,
        z_index: None,
    });

    paginator.add_fragment(fragment, measure.height, 0.0, 0.0);
}

/// Places anchored shapes at page coordinates.
fn layout_shape(block: &ShapeBlock, measure: &ShapeExtent, paginator: &mut Paginator) {
    if block.position.is_some() {
        let (x, y) = resolve_object_position(
            block.position.as_ref(),
            measure.width,
            measure.height,
            paginator,
        );
        let state_idx = paginator.get_current();
        let column_x = paginator.get_column_x(paginator.state(state_idx).column_index);
        paginator.push_fragment_direct(Fragment::Shape(ShapeFragment {
            block_id: block.id.clone(),
            wrap_offset_x: Some(x - column_x),
            x,
            y,
            width: measure.width,
            height: measure.height,
            pm_start: block.pm_start,
            pm_end: block.pm_end,
            doc_start: block.doc_start,
            doc_end: block.doc_end,
            is_anchored: Some(true),
            z_index: Some(if block.behind_doc.unwrap_or(false) {
                -1.0
            } else {
                block.relative_height.unwrap_or(1).clamp(1, 2_147_483_647) as f64
            }),
        }));
        return;
    }
    let state_idx = paginator.ensure_fits(measure.height);
    let column_index = paginator.state(state_idx).column_index;
    let fragment = Fragment::Shape(ShapeFragment {
        block_id: block.id.clone(),
        wrap_offset_x: None,
        x: paginator.get_column_x(column_index),
        y: 0.0,
        width: measure.width,
        height: measure.height,
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        doc_start: block.doc_start,
        doc_end: block.doc_end,
        is_anchored: None,
        z_index: None,
    });
    paginator.add_fragment(fragment, measure.height, 0.0, 0.0);
}

/// Charts use the same bbox placement rule as an in-flow image.
fn layout_chart(block: &ChartBlock, measure: &ChartExtent, paginator: &mut Paginator) {
    let state_idx = paginator.ensure_fits(measure.height);
    let column_index = paginator.state(state_idx).column_index;
    let fragment = Fragment::Chart(ChartFragment {
        block_id: block.id.clone(),
        x: paginator.get_column_x(column_index),
        y: 0.0,
        width: measure.width,
        height: measure.height,
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        doc_start: block.doc_start,
        doc_end: block.doc_end,
        is_anchored: None,
        z_index: None,
    });
    paginator.add_fragment(fragment, measure.height, 0.0, 0.0);
}

/// Resolves a DrawingML anchor to a page-coordinate origin.
///
/// A simple position is taken verbatim. Otherwise each axis picks a band from
/// its `relativeFrom` — page, margin box, an individual margin strip, or the
/// current column horizontally and the flow region vertically — and then
/// applies either an explicit offset or an alignment within that band. The
/// `insideMargin` and `outsideMargin` bands swap sides with page parity.
fn resolve_object_position(
    position: Option<&ImageRunPosition>,
    width: f64,
    height: f64,
    paginator: &mut Paginator,
) -> (f64, f64) {
    let state_idx = paginator.get_current();
    let state = paginator.state(state_idx);
    let page = &paginator.pages[state.page_index];
    let margins = page.body_anchor_margins.as_ref().unwrap_or(&page.margins);
    let column_x = paginator.get_column_x(state.column_index);
    crate::anchor::resolve_position(
        position,
        width,
        height,
        &crate::anchor::AnchorFrame {
            page_width: page.size.w,
            page_height: page.size.h,
            margin_left: margins.left,
            margin_right: margins.right,
            margin_top: margins.top,
            margin_bottom: margins.bottom,
            flow_x: column_x,
            flow_y: state.pen_y,
            flow_width: paginator.column_width(),
            flow_height: state.content_limit - state.pen_y,
            odd_page: page.number % 2 == 1,
        },
    )
}

/// Places an anchored image at its resolved anchor without moving the pen.
/// `behindDoc` decides whether it paints under or over body content.
fn layout_anchored_image(block: &ImageBlock, measure: &ImageExtent, paginator: &mut Paginator) {
    let anchor = block.anchor.as_ref().expect("anchored image has anchor");

    let (resolved_x, resolved_y) = resolve_object_position(
        anchor.position.as_ref(),
        measure.width,
        measure.height,
        paginator,
    );
    let x = if anchor.position.is_some() {
        resolved_x
    } else {
        anchor.offset_h.unwrap_or(resolved_x)
    };
    let mut y = if anchor.position.is_some() {
        resolved_y
    } else {
        anchor.offset_v.unwrap_or(resolved_y)
    };
    if anchor.allow_overlap == Some(false) {
        let state_idx = paginator.get_current();
        let page_index = paginator.state(state_idx).page_index;
        for existing in &paginator.pages[page_index].fragments {
            let (ex, ey, ew, eh) = match existing {
                Fragment::Paragraph(value) => (value.x, value.y, value.width, value.height),
                Fragment::Table(value) => (value.x, value.y, value.width, value.height),
                Fragment::Image(value) => (value.x, value.y, value.width, value.height),
                Fragment::Shape(value) => (value.x, value.y, value.width, value.height),
                Fragment::Chart(value) => (value.x, value.y, value.width, value.height),
                Fragment::TextBox(value) => (value.x, value.y, value.width, value.height),
            };
            if x < ex + ew && x + measure.width > ex && y < ey + eh && y + measure.height > ey {
                y = ey + eh;
            }
        }
    }

    let fragment = Fragment::Image(ImageFragment {
        block_id: block.id.clone(),
        x,
        y,
        width: measure.width,
        height: measure.height,
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        is_anchored: Some(true),
        z_index: Some(if anchor.behind_doc.unwrap_or(false) {
            -1.0
        } else {
            anchor
                .relative_height
                .or_else(|| {
                    anchor
                        .position
                        .as_ref()
                        .and_then(|value| value.relative_height)
                })
                .unwrap_or(1)
                .clamp(1, 2_147_483_647) as f64
        }),
    });

    paginator.push_fragment_direct(fragment);
}

/// Places floating text boxes as overlays and inline text boxes in flow.
fn layout_text_box(block: &TextBoxBlock, measure: &TextBoxExtent, paginator: &mut Paginator) {
    if is_floating_text_box_block(block) {
        let (x, y) = resolve_object_position(
            block.position.as_ref(),
            measure.width,
            measure.height,
            paginator,
        );
        let fragment = Fragment::TextBox(TextBoxFragment {
            block_id: block.id.clone(),
            x,
            y,
            width: measure.width,
            height: measure.height,
            pm_start: block.pm_start,
            pm_end: block.pm_end,
            is_floating: Some(true),
            z_index: Some(if block.wrap_type.as_deref() == Some("behind") {
                -1.0
            } else {
                1.0
            }),
        });
        paginator.push_fragment_direct(fragment);
        return;
    }

    let state_idx = paginator.ensure_fits(measure.height);
    let column_index = paginator.state(state_idx).column_index;

    let fragment = Fragment::TextBox(TextBoxFragment {
        block_id: block.id.clone(),
        x: paginator.get_column_x(column_index),
        y: 0.0, // set by add_fragment
        width: measure.width,
        height: measure.height,
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        is_floating: None,
        z_index: None,
    });

    paginator.add_fragment(fragment, measure.height, 0.0, 0.0);
}

#[cfg(test)]
mod pagination_rule_tests {
    use super::*;
    use serde_json::json;

    fn line(height: f64) -> serde_json::Value {
        json!({
            "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
            "width": 10, "ascent": 8, "descent": 2, "lineHeight": height,
        })
    }

    fn paragraph(
        id: u32,
        lines: usize,
        height: f64,
        attrs: serde_json::Value,
    ) -> serde_json::Value {
        json!({
            "block": {
                "kind": "paragraph", "id": id,
                "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
                "attrs": attrs,
            },
            "measure": {
                "kind": "paragraph",
                "lines": vec![line(height); lines],
                "totalHeight": lines as f64 * height,
            },
        })
    }

    fn paragraph_with_line_heights(
        id: u32,
        heights: &[f64],
        attrs: serde_json::Value,
    ) -> serde_json::Value {
        let mut block = paragraph(id, 0, 0.0, attrs);
        block["measure"]["lines"] = json!(heights.iter().copied().map(line).collect::<Vec<_>>());
        block["measure"]["totalHeight"] = json!(heights.iter().sum::<f64>());
        block
    }

    fn layout(measured: Vec<serde_json::Value>) -> Layout {
        let mut input: Input = serde_json::from_value(json!({
            "measured": measured,
            "options": {
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 },
            },
        }))
        .unwrap();
        layout_document(&mut input).unwrap()
    }

    fn input(measured: Vec<serde_json::Value>) -> Input {
        serde_json::from_value(json!({
            "measured": measured,
            "options": {
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 },
            },
        }))
        .unwrap()
    }

    fn widow_rounding_input(
        page_height: f64,
        top: f64,
        bottom: f64,
        before: f64,
        heights: &[f64],
    ) -> Input {
        let mut measured = paragraph(
            1,
            heights.len(),
            heights[0],
            json!({
                "spacing": {"before": before},
            }),
        );
        measured["measure"]["lines"] = json!(heights.iter().copied().map(line).collect::<Vec<_>>());
        serde_json::from_value(json!({
            "measured": [measured],
            "options": {
                "pageSize": {"w": 500, "h": page_height},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "sectionPageFloatBands": [{"default": [{"top": top, "bottom": bottom}]}],
            },
        }))
        .unwrap()
    }

    fn assert_all_paragraph_lines(value: &mut Input, count: usize) -> Layout {
        let result = layout_document(value).unwrap();
        let mut next_line = 0;
        for fragment in result.pages.iter().flat_map(|page| &page.fragments) {
            let Fragment::Paragraph(fragment) = fragment else {
                panic!("paragraph expected");
            };
            assert_eq!(fragment.from_line, next_line);
            assert!(fragment.to_line > fragment.from_line);
            next_line = fragment.to_line;
        }
        assert_eq!(next_line, count);
        result
    }

    #[test]
    fn widow_control_terminates_when_fractional_spacing_rounds_two_lines_to_the_gap() {
        let mut value = widow_rounding_input(300.0, 176.0, 200.0, 57.7, &[11.15; 4]);
        let result = assert_all_paragraph_lines(&mut value, 4);
        let Fragment::Paragraph(first) = &result.pages[0].fragments[0] else {
            panic!("paragraph expected");
        };
        assert_eq!(first.to_line, 2);
        value.options.section_page_float_bands = None;
        assert_eq!(assert_all_paragraph_lines(&mut value, 4).pages.len(), 1);
    }

    #[test]
    fn widow_control_terminates_when_line_rounding_matches_the_gap_with_spacing() {
        let mut value =
            widow_rounding_input(500.0, 200.0, 220.0, 64.0, &[20.0, 20.000000000000007]);
        let result = assert_all_paragraph_lines(&mut value, 2);
        assert_eq!(result.pages.len(), 1);
        assert_eq!(result.pages[0].fragments.len(), 1);
        value.options.section_page_float_bands = None;
        assert_eq!(assert_all_paragraph_lines(&mut value, 2).pages.len(), 1);
    }

    #[test]
    fn a_kept_row_pair_terminates_when_rounding_fits_it_at_the_cursor() {
        let heights = [20.0, 20.000000000000007];
        let cell = |row: usize| {
            json!({ "id": 30 + row, "blocks": [{
                "kind": "paragraph", "id": 40 + row,
                "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
                "attrs": { "keepNext": row == 0 },
            }] })
        };
        let extent = |row: usize| {
            json!({ "height": heights[row], "cells": [{ "width": 100, "height": heights[row],
                "blocks": [{ "kind": "paragraph", "lines": [line(heights[row])],
                             "totalHeight": heights[row] }] }] })
        };
        let table = json!({
            "block": {
                "kind": "table", "id": 2,
                "rows": [{ "id": 20, "cells": [cell(0)] }, { "id": 21, "cells": [cell(1)] }],
                "columnWidths": [100],
            },
            "measure": {
                "kind": "table", "columnWidths": [100],
                "totalWidth": 100, "totalHeight": heights[0] + heights[1],
                "rows": [extent(0), extent(1)],
            },
        });
        let mut value: Input = serde_json::from_value(json!({
            "measured": [paragraph(1, 1, 20.0, json!({"spacing": {"after": 64}})), table],
            "options": {
                "pageSize": {"w": 500, "h": 500},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "sectionPageFloatBands": [{"default": [{"top": 220, "bottom": 240}]}],
            },
        }))
        .unwrap();
        let result = layout_document(&mut value).unwrap();
        let last_row_end = result
            .pages
            .iter()
            .flat_map(|page| &page.fragments)
            .filter_map(|fragment| match fragment {
                Fragment::Table(table) => Some(table.row_end),
                _ => None,
            })
            .next_back();
        assert_eq!(last_row_end, Some(2));
    }

    #[test]
    fn a_table_row_taller_than_the_room_above_a_float_band_moves_below_it() {
        let table = |height: f64| {
            json!({
                "block": {
                    "kind": "table", "id": 2,
                    "rows": [{ "id": 20, "cells": [{ "id": 30, "blocks": [{
                        "kind": "paragraph", "id": 40,
                        "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
                    }] }] }],
                    "columnWidths": [100],
                },
                "measure": {
                    "kind": "table", "columnWidths": [100],
                    "totalWidth": 100, "totalHeight": height,
                    "rows": [{ "height": height, "cells": [{ "width": 100, "height": height,
                        "blocks": [{ "kind": "paragraph", "lines": [line(height)],
                                     "totalHeight": height }] }] }],
                },
            })
        };
        for (height, y) in [(40.0, Some(200.0)), (250.0, None)] {
            let mut value: Input = serde_json::from_value(json!({
                "measured": [table(height)],
                "options": {
                    "pageSize": {"w": 500, "h": 500},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                    "sectionPageFloatBands": [{"default": [{"top": 120, "bottom": 200}]}],
                },
            }))
            .unwrap();
            let result = layout_document(&mut value).unwrap();
            let fragments: Vec<_> = result
                .pages
                .iter()
                .flat_map(|page| &page.fragments)
                .collect();
            let [Fragment::Table(fragment)] = fragments.as_slice() else {
                panic!("one table fragment expected");
            };
            assert_eq!(fragment.row_end, 1, "{height}");
            if let Some(y) = y {
                assert_eq!((result.pages.len(), fragment.y), (1, y));
            }
        }
    }

    #[test]
    fn a_repeated_header_row_below_a_float_band_terminates_when_rounding_fits_the_slice() {
        let rows = [
            (true, vec![64.0]),
            (false, vec![20.0; 8]),
            (false, vec![40.000000000000007, 40.0]),
        ];
        let block_rows: Vec<_> = rows
            .iter()
            .enumerate()
            .map(|(index, (header, _))| {
                json!({ "id": 20 + index, "isHeader": header, "cells": [{ "id": 30 + index,
                    "padding": {"top": 0, "right": 0, "bottom": 0, "left": 0},
                    "blocks": [{ "kind": "paragraph", "id": 40 + index,
                        "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
                        "attrs": { "widowControl": false } }] }] })
            })
            .collect();
        let extents: Vec<_> = rows
            .iter()
            .map(|(_, lines)| {
                let height: f64 = lines.iter().sum();
                json!({ "height": height, "cells": [{ "width": 100, "height": height,
                    "blocks": [{ "kind": "paragraph",
                        "lines": lines.iter().copied().map(line).collect::<Vec<_>>(),
                        "totalHeight": height }] }] })
            })
            .collect();
        let total: f64 = rows.iter().flat_map(|(_, lines)| lines).sum();
        let mut value: Input = serde_json::from_value(json!({
            "measured": [{
                "block": { "kind": "table", "id": 2, "rows": block_rows, "columnWidths": [100] },
                "measure": { "kind": "table", "columnWidths": [100], "totalWidth": 100,
                             "totalHeight": total, "rows": extents },
            }],
            "options": {
                "pageSize": {"w": 500, "h": 580},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "sectionPageFloatBands": [{"default": [{"top": 200, "bottom": 300}]}],
            },
        }))
        .unwrap();
        let result = layout_document(&mut value).unwrap();
        let last_row_end = result
            .pages
            .iter()
            .flat_map(|page| &page.fragments)
            .filter_map(|fragment| match fragment {
                Fragment::Table(table) => Some(table.row_end),
                _ => None,
            })
            .next_back();
        assert_eq!(last_row_end, Some(3));
        assert!(result.pages.len() <= 4);
    }

    #[test]
    fn a_carried_row_slice_that_rounding_fits_between_bands_is_placed() {
        let paragraphs = [
            (vec![10.3], false),
            (vec![10.1, 25.0], true),
            (vec![20.0], false),
        ];
        let blocks: Vec<_> = paragraphs
            .iter()
            .enumerate()
            .map(|(index, (_, keep))| {
                json!({ "kind": "paragraph", "id": 40 + index,
                    "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
                    "attrs": { "keepLines": keep, "widowControl": false } })
            })
            .collect();
        let extents: Vec<_> = paragraphs
            .iter()
            .map(|(lines, _)| {
                json!({ "kind": "paragraph",
                    "lines": lines.iter().copied().map(line).collect::<Vec<_>>(),
                    "totalHeight": lines.iter().sum::<f64>() })
            })
            .collect();
        let height: f64 = paragraphs.iter().flat_map(|(lines, _)| lines).sum();
        let mut value: Input = serde_json::from_value(json!({
            "measured": [{
                "block": { "kind": "table", "id": 2, "columnWidths": [100], "rows": [
                    { "id": 20, "cells": [{ "id": 30, "blocks": blocks,
                        "padding": {"top": 0, "right": 0, "bottom": 0, "left": 0} }] }] },
                "measure": { "kind": "table", "columnWidths": [100], "totalWidth": 100,
                    "totalHeight": height, "rows": [{ "height": height, "cells": [
                        { "width": 100, "height": height, "blocks": extents }] }] },
            }],
            "options": {
                "pageSize": {"w": 500, "h": 650},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "sectionPageFloatBands": [{"default": [
                    {"top": 116, "bottom": 200}, {"top": 235.1, "bottom": 300}]}],
            },
        }))
        .unwrap();
        let result = layout_document(&mut value).unwrap();
        let fragments: Vec<_> = result
            .pages
            .iter()
            .flat_map(|page| &page.fragments)
            .filter_map(|fragment| match fragment {
                Fragment::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(fragments.last().map(|table| table.row_end), Some(1));
        assert!(result.pages.len() <= 2);
    }

    #[test]
    fn a_paragraph_split_across_columns_paints_its_float_in_each_column() {
        let floating = |height: f64, lines: usize| {
            json!({
                "block": {"kind": "paragraph", "id": format!("p{lines}"), "runs": [
                    {"kind": "text", "text": "abcdefghij"},
                    {"kind": "image", "src": "float", "width": 20, "height": 20,
                     "displayMode": "float", "wrapType": "square",
                     "position": {"vertical": {"relativeTo": "paragraph", "posOffset": 0}}},
                ], "attrs": {"widowControl": false}},
                "measure": {"kind": "paragraph", "totalHeight": height * lines as f64,
                            "lines": (0..lines).map(|index| json!({
                                "headRun": 0, "headChar": index, "tailRun": 0,
                                "tailChar": index + 1, "width": 10, "ascent": 15,
                                "descent": 5, "lineHeight": height,
                            })).collect::<Vec<_>>()},
            })
        };
        let mut value: Input = serde_json::from_value(json!({
            "measured": [paragraph(1, 5, 20.0, json!({})), floating(20.0, 15)],
            "options": {
                "pageSize": {"w": 500, "h": 492},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "columns": {"count": 2, "gap": 20},
            },
        }))
        .unwrap();
        let result = layout_document(&mut value).unwrap();
        let split: Vec<_> = result.pages[0]
            .fragments
            .iter()
            .filter_map(|fragment| match fragment {
                Fragment::Paragraph(fragment)
                    if matches!(&fragment.block_id, crate::types::BlockId::Str(id) if id == "p15") =>
                {
                    Some(fragment)
                }
                _ => None,
            })
            .collect();
        assert_eq!(split.len(), 2);
        let display: serde_json::Value = serde_json::from_str(
            &crate::display_list::build_display_list_json(
                &json!({"measured": value.measured, "options": value.options, "layout": result})
                    .to_string(),
            )
            .unwrap(),
        )
        .unwrap();
        let images = display["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|primitive| primitive["kind"] == "image")
            .count();
        assert_eq!(images, 2);
    }

    #[test]
    fn a_paragraph_crosses_two_thousand_disjoint_float_bands_on_one_page() {
        let count = 2_000;
        let bands: Vec<_> = (0..count)
            .map(|index| {
                let top = 97 + index * 2;
                json!({"top": top, "bottom": top + 1})
            })
            .collect();
        let mut value: Input = serde_json::from_value(json!({
            "measured": [paragraph(1, count + 1, 1.0, json!({"widowControl": false}))],
            "options": {
                "pageSize": {"w": 500, "h": 2 * count + 193},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "sectionPageFloatBands": [{"default": bands}],
            },
        }))
        .unwrap();
        let result = assert_all_paragraph_lines(&mut value, count + 1);
        assert_eq!(result.pages.len(), 1);
        assert_eq!(result.pages[0].fragments.len(), count + 1);
        for (index, fragment) in result.pages[0].fragments.iter().enumerate() {
            let Fragment::Paragraph(fragment) = fragment else {
                panic!("paragraph expected");
            };
            assert_eq!(fragment.y, (96 + index * 2) as f64);
            assert_eq!(fragment.height, 1.0);
        }
    }

    #[test]
    fn internal_float_bands_emit_paragraph_images_once_per_page_in_both_passes() {
        for wrap in ["square", "behind"] {
            let mut value: Input = serde_json::from_value(json!({
                "measured": [{
                    "block": {"kind": "paragraph", "id": "body", "runs": [
                        {"kind": "text", "text": "abcdefghij"},
                        {"kind": "image", "src": "float", "width": 20, "height": 20,
                         "displayMode": "float", "wrapType": wrap,
                         "position": {"vertical": {"relativeTo": "paragraph", "posOffset": 0}}},
                    ]},
                    "measure": {"kind": "paragraph", "totalHeight": 200,
                                "lines": (0..10).map(|index| json!({
                                    "headRun": 0, "headChar": index, "tailRun": 0,
                                    "tailChar": index + 1, "width": 10, "ascent": 15,
                                    "descent": 5, "lineHeight": 20,
                                })).collect::<Vec<_>>()},
                }],
                "options": {
                    "pageSize": {"w": 500, "h": 500},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                    "sectionPageFloatBands": [{"default": [{"top": 200, "bottom": 240}]}],
                },
            }))
            .unwrap();
            let result = layout_document(&mut value).unwrap();
            assert_eq!(result.pages.len(), 1);
            let fragments = &result.pages[0].fragments;
            assert_eq!(fragments.len(), 2);
            for (fragment, expected_y) in fragments.iter().zip([96.0, 240.0]) {
                let Fragment::Paragraph(fragment) = fragment else {
                    panic!("paragraph expected");
                };
                assert_eq!(fragment.y, expected_y);
            }
            let display: serde_json::Value = serde_json::from_str(
                &crate::display_list::build_display_list_json(
                    &json!({"measured": value.measured, "options": value.options,
                            "layout": result})
                    .to_string(),
                )
                .unwrap(),
            )
            .unwrap();
            let primitives = display["pages"][0]["primitives"].as_array().unwrap();
            let images: Vec<_> = primitives
                .iter()
                .enumerate()
                .filter(|(_, primitive)| primitive["kind"] == "image")
                .collect();
            assert_eq!(images.len(), 1);
            assert_eq!(images[0].1["y"], 96);
            for (index, primitive) in primitives.iter().enumerate() {
                if primitive["kind"] == "text" {
                    assert_eq!(images[0].0 < index, wrap == "behind");
                }
            }
        }
    }

    #[test]
    fn balanced_columns_with_an_internal_float_band_stay_on_one_page() {
        let mut value = input(vec![paragraph(1, 16, 20.0, json!({}))]);
        value.options.page_size = Some(crate::types::Size { w: 500.0, h: 300.0 });
        value.options.columns = Some(
            serde_json::from_value(json!({
                "count": 2, "gap": 20,
            }))
            .unwrap(),
        );
        value.options.section_page_float_bands = Some(
            serde_json::from_value(json!([{
                "default": [{"top": 70, "bottom": 110}],
            }]))
            .unwrap(),
        );

        let result = layout_document(&mut value).unwrap();

        assert_eq!(result.pages.len(), 1);
        let fragments: Vec<_> = result.pages[0]
            .fragments
            .iter()
            .filter_map(|fragment| match fragment {
                Fragment::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .collect();
        assert_eq!(
            fragments
                .iter()
                .map(|p| p.to_line - p.from_line)
                .sum::<usize>(),
            16
        );
        assert!(fragments.iter().any(|p| p.x != fragments[0].x));
        assert!(
            fragments
                .iter()
                .all(|p| p.y + p.height <= 70.0 || p.y >= 110.0)
        );
    }

    #[test]
    fn a_standalone_break_paragraph_suppresses_leading_spacing_but_a_column_break_preserves_it() {
        let mut value = input(vec![
            paragraph(1, 1, 10.0, json!({})),
            json!({"block":{"kind":"pageBreak","id":"page"},"measure":{"kind":"pageBreak"}}),
            paragraph(2, 1, 10.0, json!({"spacing":{"before":20}})),
            json!({"block":{"kind":"columnBreak","id":"column"},"measure":{"kind":"columnBreak"}}),
            paragraph(3, 1, 10.0, json!({"spacing":{"before":20}})),
        ]);
        let result = layout_document(&mut value).unwrap();
        assert_eq!(result.pages.len(), 3);
        let Fragment::Paragraph(after_page) = &result.pages[1].fragments[0] else {
            panic!()
        };
        let Fragment::Paragraph(after_column) = &result.pages[2].fragments[0] else {
            panic!()
        };
        assert_eq!(after_page.y, 10.0);
        assert_eq!(after_column.y, 30.0);
        let recorded = layout_document_checkpointed(&mut value).unwrap();
        let checkpoint = recorded
            .checkpoints
            .iter()
            .find(|checkpoint| checkpoint.page_index == 1)
            .unwrap();
        assert!(checkpoint.flow.leading_spacing_spent.is_infinite());
    }

    /// Measured against Word 16.113 (oxi-ja-policies-01 pages 40 and 45, and
    /// hand-authored probes): a paragraph that breaks the page itself, by
    /// `w:br w:type="page"` or by `w:pageBreakBefore`, keeps its space-before
    /// on the new page, however full the previous page was.
    #[test]
    fn a_paragraph_that_breaks_its_own_page_keeps_leading_spacing() {
        for attrs in [
            json!({"spacing":{"before":20},"pageBreakBeforeRun":true}),
            json!({"spacing":{"before":20},"pageBreakBefore":true}),
        ] {
            for filler in [1, 5, 9] {
                let mut measured = vec![paragraph(0, filler, 10.0, json!({}))];
                measured.push(paragraph(1, 1, 10.0, attrs.clone()));
                let mut value = input(measured);
                let result = layout_document(&mut value).unwrap();
                assert_eq!(result.pages.len(), 2, "filler {filler}");
                let Fragment::Paragraph(after_break) = &result.pages[1].fragments[0] else {
                    panic!()
                };
                assert_eq!(after_break.y, 30.0, "filler {filler}");
                let recorded = layout_document_checkpointed(&mut value).unwrap();
                let checkpoint = recorded
                    .checkpoints
                    .iter()
                    .find(|checkpoint| checkpoint.page_index == 1)
                    .unwrap();
                assert_eq!(
                    checkpoint.flow.leading_spacing_spent, 0.0,
                    "filler {filler}"
                );
            }
        }
    }

    #[test]
    fn an_automatic_page_break_discards_leading_spacing() {
        for attrs in [
            json!({"spacing":{"before":20}}),
            json!({"spacing":{"before":20},"keepNext":true}),
            json!({"spacing":{"before":20},"keepLines":true}),
        ] {
            let mut value = input(vec![
                paragraph(1, 1, 85.0, json!({"spacing":{"after":30}})),
                paragraph(2, 1, 20.0, attrs),
                paragraph(3, 1, 20.0, json!({})),
            ]);
            let result = layout_document(&mut value).unwrap();
            assert_eq!(result.pages.len(), 2);
            let Fragment::Paragraph(heading) = &result.pages[1].fragments[0] else {
                panic!()
            };
            let Fragment::Paragraph(body) = &result.pages[1].fragments[1] else {
                panic!()
            };
            assert_eq!(heading.y, 10.0);
            assert_eq!(body.y, 30.0);
        }
    }

    /// Measured against Word 16.113 (hand-authored probes, prev space-after
    /// 0/12/24/36pt against 6/24/48pt space-before): the collapsed gap is
    /// spent from the bottom up, so an authored break carries only
    /// `max(0, before - after)` onto the new page.
    #[test]
    fn an_authored_break_spends_the_previous_space_after() {
        for (after, before, expected) in [
            (0.0, 24.0, 34.0),
            (12.0, 24.0, 22.0),
            (24.0, 24.0, 10.0),
            (36.0, 24.0, 10.0),
            (36.0, 48.0, 22.0),
        ] {
            for attrs in [
                json!({"spacing":{"before":before},"pageBreakBefore":true}),
                json!({"spacing":{"before":before},"pageBreakBeforeRun":true}),
            ] {
                let mut value = input(vec![
                    paragraph(1, 1, 10.0, json!({"spacing":{"after":after}})),
                    paragraph(2, 1, 10.0, attrs),
                ]);
                let result = layout_document(&mut value).unwrap();
                assert_eq!(result.pages.len(), 2, "after {after} before {before}");
                let Fragment::Paragraph(after_break) = &result.pages[1].fragments[0] else {
                    panic!()
                };
                assert_eq!(after_break.y, expected, "after {after} before {before}");
            }
        }
    }

    /// Measured against Word 16.113: `w:contextualSpacing` between same-style
    /// neighbours zeroes the space-before before pagination, so an authored
    /// break has nothing to carry; a different previous style leaves it whole.
    #[test]
    fn contextual_spacing_leaves_an_authored_break_nothing_to_carry() {
        let target = json!({
            "styleId": "List", "effectiveStyleId": "List", "contextualSpacing": true,
            "spacing": {"before": 20}, "pageBreakBefore": true
        });
        for (previous, expected) in [
            (
                json!({"styleId": "List", "effectiveStyleId": "List", "contextualSpacing": true}),
                10.0,
            ),
            (json!({"styleId": "Body", "effectiveStyleId": "Body"}), 30.0),
        ] {
            let mut value = input(vec![
                paragraph(1, 1, 10.0, previous),
                paragraph(2, 1, 10.0, target.clone()),
            ]);
            let result = layout_document(&mut value).unwrap();
            assert_eq!(result.pages.len(), 2);
            let Fragment::Paragraph(after_break) = &result.pages[1].fragments[0] else {
                panic!()
            };
            assert_eq!(after_break.y, expected);
        }
    }

    #[test]
    fn anchored_shape_does_not_advance_body_flow() {
        for (anchored, wrap, overlay) in [
            (false, "none", false),
            (true, "none", true),
            (true, "square", true),
            (true, "tight", true),
            (true, "through", true),
            (true, "topAndBottom", true),
        ] {
            let mut shape = json!({
                "kind": "shape", "id": "shape", "shapeType": "rect",
                "width": 50, "height": 40, "geometryPath": [], "children": [],
                "wrapType": wrap
            });
            if anchored {
                shape["position"] = json!({
                    "horizontal": {"relativeTo": "page", "posOffset": 100},
                    "vertical": {"relativeTo": "page", "posOffset": 10}
                });
            }
            let mut input: Input = serde_json::from_value(json!({
                "measured": [
                    {"block": shape, "measure": {"kind": "shape", "width": 50, "height": 40}},
                    {"block": {"kind": "image", "id": "body", "src": "", "width": 20, "height": 20},
                     "measure": {"kind": "image", "width": 20, "height": 20}}
                ],
                "options": {"pageSize": {"w": 300, "h": 200}, "margins": {"left": 20, "right": 20, "top": 20, "bottom": 20}}
            })).unwrap();
            let layout = layout_document(&mut input).unwrap();
            assert_eq!(layout.pages.len(), 1);
            let Fragment::Shape(shape) = &layout.pages[0].fragments[0] else {
                panic!("shape expected")
            };
            let Fragment::Image(body) = &layout.pages[0].fragments[1] else {
                panic!("image expected")
            };
            assert_eq!(body.y, if overlay { 20.0 } else { 60.0 });
            assert_eq!(
                (shape.x, shape.y),
                if overlay { (100.0, 10.0) } else { (20.0, 20.0) }
            );
        }
    }

    #[test]
    fn incremental_layout_stops_at_converged_page_start() {
        let measured: Vec<_> = (0..15)
            .map(|id| paragraph(id, 1, 20.0, json!({})))
            .collect();
        let mut previous_input = input(measured.clone());
        let previous = layout_document_checkpointed(&mut previous_input).unwrap();
        assert!(previous.checkpoints.len() >= 3);

        let mut incremental_input = input(measured.clone());
        let previous_fingerprints = vec![1_u64; measured.len()];
        let mut next_fingerprints = previous_fingerprints.clone();
        next_fingerprints[0] = 2;
        let mut previous_layout = previous.layout;
        let incremental = layout_document_incremental(
            &mut incremental_input,
            &mut previous_layout,
            &previous.checkpoints,
            &previous_fingerprints,
            &next_fingerprints,
            0,
        )
        .unwrap();

        let mut full_input = input(measured);
        let full = layout_document_checkpointed(&mut full_input).unwrap();
        assert_eq!(
            serde_json::to_string(&incremental.layout).unwrap(),
            serde_json::to_string(&full.layout).unwrap()
        );
        assert!(incremental.placed_blocks < full.placed_blocks);
        assert_eq!(incremental.rebuilt_page_start, 0);
        assert_eq!(incremental.rebuilt_page_end, 1);
    }

    #[test]
    fn incremental_resumes_when_note_reservations_change_after_the_checkpoint() {
        let measured: Vec<_> = (0..30)
            .map(|id| paragraph(id, 1, 20.0, json!({})))
            .collect();
        for (previous_reservations, next_reservations) in [
            (json!({"4": 20}), json!({"3": 20})),
            (json!({"3": 20}), json!({"4": 20})),
            (json!({"4": 20}), json!({"1": 0, "2": 0, "3": 20})),
            (json!({"1": 0, "2": 0, "3": 20}), json!({"4": 20})),
        ] {
            let mut previous_input = input(measured.clone());
            previous_input.options.footnote_reserved_heights =
                serde_json::from_value(previous_reservations).unwrap();
            let previous = layout_document_checkpointed(&mut previous_input).unwrap();
            let mut next = measured.clone();
            next[11] = paragraph(11, 1, 40.0, json!({}));
            let mut next_input = input(next);
            next_input.options.footnote_reserved_heights =
                serde_json::from_value(next_reservations).unwrap();
            let previous_fingerprints = vec![1_u64; measured.len()];
            let mut next_fingerprints = previous_fingerprints.clone();
            next_fingerprints[11] = 2;
            let incremental = layout_document_incremental(
                &mut next_input.clone(),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                10,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut next_input).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 1);
            assert_eq!(incremental.rebuilt_page_end, full.layout.pages.len());
            assert_eq!(incremental.placed_blocks, measured.len() - 5);
            assert!(incremental.placed_blocks < full.placed_blocks);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_converges_after_changed_note_reservations() {
        let measured: Vec<_> = (0..28)
            .map(|id| {
                paragraph(
                    id,
                    1,
                    20.0,
                    json!({"pageBreakBefore": id > 0 && id % 4 == 0}),
                )
            })
            .collect();
        for (previous_reservations, next_reservations) in [
            (json!({"3": 10}), json!({"4": 10, "5": 0, "40": 0})),
            (json!({"4": 10, "5": 0, "40": 0}), json!({"3": 10})),
        ] {
            let mut previous_input = input(measured.clone());
            previous_input.options.footnote_reserved_heights =
                serde_json::from_value(previous_reservations).unwrap();
            let previous = layout_document_checkpointed(&mut previous_input).unwrap();
            let mut next = measured.clone();
            next[9]["block"]["runs"][0]["text"] = json!("y");
            let mut next_input = input(next);
            next_input.options.footnote_reserved_heights =
                serde_json::from_value(next_reservations).unwrap();
            let previous_fingerprints = vec![1_u64; measured.len()];
            let mut next_fingerprints = previous_fingerprints.clone();
            next_fingerprints[9] = 2;
            let incremental = layout_document_incremental(
                &mut next_input.clone(),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                8,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut next_input).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 1);
            assert_eq!(incremental.rebuilt_page_end, 4);
            assert!(incremental.placed_blocks < full.placed_blocks);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_does_not_converge_before_a_changed_suffix_reservation() {
        let measured: Vec<_> = (0..28)
            .map(|id| {
                paragraph(
                    id,
                    1,
                    20.0,
                    json!({"pageBreakBefore": id > 0 && id % 4 == 0}),
                )
            })
            .collect();
        for changed_page in [5_u32, 6] {
            let mut previous_input = input(measured.clone());
            previous_input.options.footnote_reserved_heights =
                serde_json::from_value(json!({"3": 10})).unwrap();
            let previous = layout_document_checkpointed(&mut previous_input).unwrap();
            let mut next = measured.clone();
            next[9]["block"]["runs"][0]["text"] = json!("y");
            let mut next_input = input(next);
            next_input.options.footnote_reserved_heights =
                serde_json::from_value(json!({"4": 10})).unwrap();
            next_input
                .options
                .footnote_reserved_heights
                .as_mut()
                .unwrap()
                .insert(changed_page.to_string(), 10.0);
            let previous_fingerprints = vec![1_u64; measured.len()];
            let mut next_fingerprints = previous_fingerprints.clone();
            next_fingerprints[9] = 2;
            let incremental = layout_document_incremental(
                &mut next_input.clone(),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                8,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut next_input).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 1);
            assert_eq!(incremental.rebuilt_page_end, changed_page as usize);
            assert!(incremental.placed_blocks < full.placed_blocks);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_refuses_note_reservation_changes_at_or_before_the_checkpoint() {
        let measured: Vec<_> = (0..30)
            .map(|id| paragraph(id, 1, 20.0, json!({})))
            .collect();
        let mut previous_input = input(measured.clone());
        previous_input.options.footnote_reserved_heights =
            serde_json::from_value(json!({"4": 20})).unwrap();
        let previous = layout_document_checkpointed(&mut previous_input).unwrap();
        for changed_page in [1_u32, 2] {
            let mut next = measured.clone();
            next[11] = paragraph(11, 1, 40.0, json!({}));
            let mut next_input = input(next);
            let mut reservations = previous_input.options.footnote_reserved_heights.clone();
            reservations
                .as_mut()
                .unwrap()
                .insert(changed_page.to_string(), 20.0);
            next_input.options.footnote_reserved_heights = reservations;
            let previous_fingerprints = vec![1_u64; measured.len()];
            let mut next_fingerprints = previous_fingerprints.clone();
            next_fingerprints[11] = 2;
            let mut retained = previous.layout.clone();
            let result = layout_document_incremental(
                &mut next_input.clone(),
                &mut retained,
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                10,
            );
            assert!(matches!(
                result,
                Err(LayoutError::Unsupported(message))
                    if message == "checkpoint note reservations changed"
            ));
            assert_eq!(
                serde_json::to_vec(&retained).unwrap(),
                serde_json::to_vec(&previous.layout).unwrap()
            );
            let incremental = layout_document_incremental(
                &mut next_input.clone(),
                &mut retained,
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                (changed_page as usize - 1) * 5,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut next_input).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 0);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_restarts_around_split_paragraphs_match_full() {
        for (preceding, lines, height, widow, keep_lines, slices) in [
            (95.0, 2, 10.0, false, false, vec![(3, 0, 1), (3, 1, 2)]),
            (90.0, 2, 10.0, false, false, vec![(2, 0, 1), (3, 1, 2)]),
            (40.0, 4, 20.0, true, false, vec![(2, 0, 2), (3, 2, 4)]),
            (70.0, 4, 20.0, true, false, vec![(3, 0, 4)]),
            (95.0, 2, 10.0, true, false, vec![(3, 0, 2)]),
            (95.0, 2, 10.0, false, true, vec![(3, 0, 2)]),
        ] {
            for after_split in [false, true] {
                let blocks = |edited: bool| {
                    let mut blocks = vec![
                        paragraph(0, 1, 100.0, json!({})),
                        paragraph(1, 1, 100.0, json!({})),
                        paragraph(2, 1, preceding, json!({})),
                        paragraph(
                            3,
                            lines,
                            height,
                            json!({
                                "widowControl": widow, "keepLines": keep_lines,
                            }),
                        ),
                    ];
                    if after_split {
                        blocks.push(paragraph(4, 1, 100.0, json!({})));
                    }
                    blocks.push(paragraph(5, 1, if edited { 20.0 } else { 10.0 }, json!({})));
                    blocks.extend((6..12).map(|id| paragraph(id, 1, 100.0, json!({}))));
                    blocks
                };
                let mut previous_input = input(blocks(false));
                let previous = layout_document_checkpointed(&mut previous_input).unwrap();
                assert_eq!(paragraph_slices(&previous.layout, 3.0), slices);
                let dirty = if after_split { 5 } else { 4 };
                let fingerprints = vec![1_u64; previous_input.measured.len()];
                let mut next_fingerprints = fingerprints.clone();
                next_fingerprints[dirty] = 2;
                let next = blocks(true);
                let incremental = layout_document_incremental(
                    &mut input(next.clone()),
                    &mut previous.layout.clone(),
                    &previous.checkpoints,
                    &fingerprints,
                    &next_fingerprints,
                    dirty,
                )
                .unwrap();
                let full = layout_document_checkpointed(&mut input(next)).unwrap();
                let restart_page = if after_split {
                    4
                } else if slices.len() == 1 {
                    3
                } else {
                    2
                };
                assert_eq!(incremental.rebuilt_page_start, restart_page);
                assert!(incremental.placed_blocks < full.placed_blocks);
                assert_eq!(
                    serde_json::to_vec(&incremental.layout).unwrap(),
                    serde_json::to_vec(&full.layout).unwrap()
                );
                assert_eq!(incremental.checkpoints, full.checkpoints);
            }
        }
    }

    #[test]
    fn incremental_resumes_at_pristine_multiple_fragment_paragraphs() {
        for repetitions in [1, 2] {
            let blocks = |edited: bool| {
                let mut blocks = vec![paragraph(0, 1, 10.0, json!({}))];
                for repeat in 0..repetitions {
                    let id = blocks.len() as u32;
                    blocks.push(json!({
                        "block": {"kind": "pageBreak", "id": id},
                        "measure": {"kind": "pageBreak"},
                    }));
                    blocks.push(paragraph_with_line_heights(
                        id + 1,
                        &[20.0, 20.0, 20.0, 90.0],
                        json!({"widowControl": true}),
                    ));
                    blocks.push(paragraph(
                        id + 2,
                        1,
                        if edited && repeat + 1 == repetitions {
                            20.0
                        } else {
                            10.0
                        },
                        json!({}),
                    ));
                }
                let id = blocks.len() as u32;
                blocks.push(json!({
                    "block": {"kind": "pageBreak", "id": id},
                    "measure": {"kind": "pageBreak"},
                }));
                blocks.push(paragraph(id + 1, 1, 100.0, json!({})));
                blocks
            };
            let mut previous_input = input(blocks(false));
            let previous = layout_document_checkpointed(&mut previous_input).unwrap();
            for repeat in 0..repetitions {
                let opening_page = 1 + repeat * 2;
                assert_eq!(
                    paragraph_slices(&previous.layout, (2 + repeat * 3) as f64),
                    vec![
                        (opening_page, 0, 2),
                        (opening_page, 2, 3),
                        (opening_page + 1, 3, 4),
                    ]
                );
            }
            let dirty = repetitions * 3;
            let fingerprints = vec![1_u64; previous_input.measured.len()];
            let mut next_fingerprints = fingerprints.clone();
            next_fingerprints[dirty] = 2;
            let next = blocks(true);
            let incremental = layout_document_incremental(
                &mut input(next.clone()),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &fingerprints,
                &next_fingerprints,
                dirty,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut input(next)).unwrap();
            assert_eq!(incremental.rebuilt_page_start, repetitions * 2 - 1);
            assert!(incremental.placed_blocks < full.placed_blocks);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn a_paragraph_spanning_pages_can_resume_at_its_clean_opening() {
        let blocks = |edited: bool| {
            vec![
                paragraph(0, 1, 100.0, json!({})),
                paragraph(1, 1, 100.0, json!({})),
                json!({"block": {"kind": "pageBreak", "id": 2}, "measure": {"kind": "pageBreak"}}),
                paragraph(3, 12, 10.0, json!({"widowControl": false})),
                paragraph(4, 1, if edited { 20.0 } else { 10.0 }, json!({})),
                paragraph(5, 1, 100.0, json!({})),
            ]
        };
        let previous = layout_document_checkpointed(&mut input(blocks(false))).unwrap();
        assert_eq!(
            paragraph_slices(&previous.layout, 3.0),
            vec![(2, 0, 10), (3, 10, 12)]
        );
        let incremental = layout_document_incremental(
            &mut input(blocks(true)),
            &mut previous.layout.clone(),
            &previous.checkpoints,
            &[1_u64; 6],
            &[1, 1, 1, 1, 2, 1],
            4,
        )
        .unwrap();
        let full = layout_document_checkpointed(&mut input(blocks(true))).unwrap();
        assert_eq!(incremental.rebuilt_page_start, 2);
        assert!(incremental.placed_blocks < full.placed_blocks);
        assert_eq!(
            serde_json::to_vec(&incremental.layout).unwrap(),
            serde_json::to_vec(&full.layout).unwrap()
        );
        assert_eq!(incremental.checkpoints, full.checkpoints);
    }

    #[test]
    fn paragraph_checkpoints_reject_multiple_opening_fragments_and_continuations() {
        for preceding in [95.0, 90.0] {
            let mut value = input(vec![
                paragraph(0, 1, 100.0, json!({})),
                paragraph(1, 1, preceding, json!({})),
                paragraph(2, 2, 10.0, json!({ "widowControl": false })),
            ]);
            let full = layout_document_checkpointed(&mut value).unwrap();
            assert!(
                !full.checkpoints.iter().any(|checkpoint| {
                    checkpoint.block_index == 2 && checkpoint.page_index == 2
                })
            );
            let mut checkpoint = full.checkpoints[1].clone();
            checkpoint.block_index = 2;
            checkpoint.page_index = 2;
            checkpoint.page_number = full.layout.pages[2].number;
            assert!(!resumable(&checkpoint, &value.measured, &full.layout.pages));
        }
    }

    #[test]
    fn incremental_convergence_matches_the_opening_fragments_fitting_geometry() {
        let value = |margin: f64| {
            let mut value = input(vec![
                paragraph(0, 10, 10.0, json!({})),
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": "continuous",
                        "margins": { "top": 10, "right": margin, "bottom": 10, "left": margin },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
                paragraph(2, 1, 10.0, json!({ "alignment": "center" })),
                paragraph(3, 1, 10.0, json!({})),
            ]);
            value.options.final_margins = serde_json::from_value(json!({
                "top": 10, "right": 30, "bottom": 10, "left": 30,
            }))
            .unwrap();
            value
        };
        let previous = layout_document_checkpointed(&mut value(10.0)).unwrap();
        let Fragment::Paragraph(previous_fragment) = &previous.layout.pages[1].fragments[0] else {
            panic!("paragraph expected")
        };
        assert_eq!(previous_fragment.width, 180.0);
        let incremental = layout_document_incremental(
            &mut value(20.0),
            &mut previous.layout.clone(),
            &previous.checkpoints,
            &[1_u64; 4],
            &[1, 2, 1, 1],
            1,
        )
        .unwrap();
        let full = layout_document_checkpointed(&mut value(20.0)).unwrap();
        assert_eq!(full.layout.pages.len(), 2);
        let Fragment::Paragraph(full_fragment) = &full.layout.pages[1].fragments[0] else {
            panic!("paragraph expected")
        };
        let Fragment::Paragraph(incremental_fragment) = &incremental.layout.pages[1].fragments[0]
        else {
            panic!("paragraph expected")
        };
        assert_eq!(full_fragment.width, 160.0);
        assert_eq!(incremental_fragment.width, full_fragment.width);
        assert_eq!(
            serde_json::to_vec(&incremental.layout).unwrap(),
            serde_json::to_vec(&full.layout).unwrap()
        );
        assert_eq!(incremental.checkpoints, full.checkpoints);
    }

    #[test]
    fn incremental_convergence_matches_fitting_geometry_with_first_page_margins() {
        let value = |margin: f64| {
            let mut value = input(vec![
                paragraph(0, 10, 10.0, json!({})),
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": "continuous",
                        "margins": { "top": 10, "right": margin, "bottom": 10, "left": margin },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
                paragraph(2, 1, 10.0, json!({ "alignment": "center" })),
                paragraph(3, 1, 10.0, json!({})),
            ]);
            value.options.final_margins = serde_json::from_value(json!({
                "top": 10, "right": 30, "bottom": 10, "left": 30,
            }))
            .unwrap();
            value.options.section_page_margins = Some(
                serde_json::from_value(json!([
                    { "first": { "top": 10, "right": 10, "bottom": 10, "left": 10 } },
                    {},
                ]))
                .unwrap(),
            );
            value
        };
        let previous = layout_document_checkpointed(&mut value(10.0)).unwrap();
        let Fragment::Paragraph(previous_fragment) = &previous.layout.pages[1].fragments[0] else {
            panic!("paragraph expected")
        };
        assert_eq!(previous_fragment.width, 180.0);
        let incremental = layout_document_incremental(
            &mut value(20.0),
            &mut previous.layout.clone(),
            &previous.checkpoints,
            &[1_u64; 4],
            &[1, 2, 1, 1],
            1,
        )
        .unwrap();
        let mut full_input = value(20.0);
        let full = layout_document_checkpointed(&mut full_input).unwrap();
        assert_eq!(full.layout.pages.len(), 2);
        assert_eq!(
            previous.layout.pages[0].margins,
            full.layout.pages[0].margins
        );
        assert_ne!(
            previous.layout.pages[1].opening_fragment_geometry,
            full.layout.pages[1].opening_fragment_geometry
        );
        let Fragment::Paragraph(full_fragment) = &full.layout.pages[1].fragments[0] else {
            panic!("paragraph expected")
        };
        let Fragment::Paragraph(incremental_fragment) = &incremental.layout.pages[1].fragments[0]
        else {
            panic!("paragraph expected")
        };
        assert_eq!(full_fragment.width, 160.0);
        assert_eq!(incremental_fragment.width, full_fragment.width);
        assert_eq!(
            serde_json::to_vec(&incremental.layout).unwrap(),
            serde_json::to_vec(&full.layout).unwrap()
        );
        assert_eq!(incremental.checkpoints, full.checkpoints);
        assert_eq!(
            incremental.layout.pages[1].opening_fragment_geometry,
            full.layout.pages[1].opening_fragment_geometry
        );
        let checkpoint = full
            .checkpoints
            .iter()
            .find(|checkpoint| checkpoint.block_index == 2 && checkpoint.page_index == 1)
            .unwrap();
        assert!(!resumable(
            checkpoint,
            &full_input.measured,
            &full.layout.pages
        ));
    }

    #[test]
    fn incremental_text_edits_keep_the_same_convergence_page() {
        for page_break_before in [false, true] {
            let measured: Vec<_> = (0..15)
                .map(|id| {
                    paragraph(
                        id,
                        1,
                        20.0,
                        json!({ "pageBreakBefore": page_break_before && id > 0 && id % 5 == 0 }),
                    )
                })
                .collect();
            let previous = layout_document_checkpointed(&mut input(measured.clone())).unwrap();
            let mut next = measured;
            next[2]["block"]["runs"][0]["text"] = json!("y");
            let mut next_fingerprints = [1_u64; 15];
            next_fingerprints[2] = 2;
            let incremental = layout_document_incremental(
                &mut input(next.clone()),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &[1_u64; 15],
                &next_fingerprints,
                2,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut input(next)).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 0);
            assert_eq!(incremental.rebuilt_page_end, 1);
            assert_eq!(incremental.placed_blocks, 6);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_text_edits_converge_across_sections_with_the_same_geometry() {
        for (break_type, placed_blocks) in [("continuous", 7), ("nextPage", 6)] {
            let mut measured: Vec<_> = (0..15)
                .map(|id| paragraph(id, 1, 20.0, json!({})))
                .collect();
            measured.insert(
                5,
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": break_type,
                        "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
            );
            let previous = layout_document_checkpointed(&mut input(measured.clone())).unwrap();
            assert_eq!(previous.layout.pages.len(), 3);
            assert_eq!(
                (
                    previous.checkpoints[1].block_index,
                    previous.checkpoints[1].section_index,
                    previous.checkpoints[1].page_index,
                ),
                (6, 1, 1)
            );
            let mut next = measured;
            next[2]["block"]["runs"][0]["text"] = json!("y");
            let mut next_fingerprints = [1_u64; 16];
            next_fingerprints[2] = 2;
            let incremental = layout_document_incremental(
                &mut input(next.clone()),
                &mut previous.layout.clone(),
                &previous.checkpoints,
                &[1_u64; 16],
                &next_fingerprints,
                2,
            )
            .unwrap();
            let full = layout_document_checkpointed(&mut input(next)).unwrap();
            assert_eq!(incremental.rebuilt_page_start, 0);
            assert_eq!(incremental.rebuilt_page_end, 1);
            assert_eq!(incremental.placed_blocks, placed_blocks);
            assert_eq!(
                serde_json::to_vec(&incremental.layout).unwrap(),
                serde_json::to_vec(&full.layout).unwrap()
            );
            assert_eq!(incremental.checkpoints, full.checkpoints);
        }
    }

    #[test]
    fn incremental_layout_skips_the_clean_pages_between_changes() {
        let keep_next = 41;
        let measured = |lines: &[usize]| -> Vec<serde_json::Value> {
            lines
                .iter()
                .enumerate()
                .map(|(id, &count)| {
                    let attrs = if id == keep_next {
                        json!({ "keepNext": true })
                    } else {
                        json!({})
                    };
                    paragraph(id as u32, count, 20.0, attrs)
                })
                .collect()
        };
        let base: Vec<usize> = (0..60).map(|id| if id == 12 { 2 } else { 1 }).collect();
        let mut previous_input = input(measured(&base));
        let previous = layout_document_checkpointed(&mut previous_input).unwrap();
        let previous_fingerprints = vec![1_u64; base.len()];

        // Each change is a block and its new line count; a same count changes only text.
        let cases: [&[(usize, usize)]; 5] = [
            &[(7, 1), (42, 1)],
            &[(7, 1), (42, 2)],
            &[(7, 2), (42, 1)],
            &[(7, 2), (12, 1), (42, 1), (50, 1)],
            &[(3, 1), (20, 2), (26, 1), (41, 2), (57, 1)],
        ];
        for changes in cases {
            let mut lines = base.clone();
            let mut next_fingerprints = previous_fingerprints.clone();
            for &(index, count) in changes {
                lines[index] = count;
                next_fingerprints[index] = 2;
            }
            let mut previous_layout = previous.layout.clone();
            let mut incremental_input = input(measured(&lines));
            let incremental = layout_document_incremental(
                &mut incremental_input,
                &mut previous_layout,
                &previous.checkpoints,
                &previous_fingerprints,
                &next_fingerprints,
                changes[0].0,
            )
            .unwrap();
            let mut full_input = input(measured(&lines));
            let full = layout_document_checkpointed(&mut full_input).unwrap();
            assert_eq!(
                serde_json::to_string(&incremental.layout).unwrap(),
                serde_json::to_string(&full.layout).unwrap(),
                "{changes:?}"
            );
            assert_eq!(incremental.checkpoints, full.checkpoints, "{changes:?}");
        }

        let mut previous_layout = previous.layout.clone();
        let mut next_fingerprints = previous_fingerprints.clone();
        next_fingerprints[7] = 2;
        next_fingerprints[42] = 2;
        let IncrementalLayout {
            checkpointed: incremental,
            rebuilt_page_ranges,
        } = layout_document_incremental_ranges(
            &mut input(measured(&base)),
            &mut previous_layout,
            &previous.checkpoints,
            &previous_fingerprints,
            &next_fingerprints,
            7,
        )
        .unwrap();
        assert_eq!(rebuilt_page_ranges, vec![1..2, 8..9]);
        assert_eq!(
            (incremental.rebuilt_page_start, incremental.rebuilt_page_end),
            (1, 9)
        );
        assert!(incremental.placed_blocks * 4 < base.len());
    }

    #[test]
    fn a_checkpoint_after_a_floating_table_keeps_the_spacing_its_page_starts_with() {
        let floating_table = json!({
            "block": {
                "kind": "table", "id": 90,
                "rows": [{ "id": 91, "cells": [{ "id": 92, "blocks": [] }] }],
                "columnWidths": [50],
                "floating": {},
            },
            "measure": {
                "kind": "table", "columnWidths": [50],
                "totalWidth": 50, "totalHeight": 20,
                "rows": [{ "height": 20, "cells": [
                    { "width": 50, "height": 20, "blocks": [] }
                ] }],
            },
        });
        // The table overflows onto the second page, whose automatic break spends
        // the following paragraph's space before.
        let measured = || {
            vec![
                paragraph(0, 9, 10.0, json!({})),
                floating_table.clone(),
                paragraph(1, 1, 10.0, json!({ "spacing": { "before": 20 } })),
            ]
        };
        let previous = layout_document_checkpointed(&mut input(measured())).unwrap();
        let table_page = previous
            .checkpoints
            .iter()
            .find(|checkpoint| checkpoint.block_index == 1)
            .unwrap();
        assert_eq!(table_page.page_index, 1);
        assert_eq!(table_page.flow.leading_spacing_spent, f64::INFINITY);
        let mut previous_layout = previous.layout.clone();
        let incremental = layout_document_incremental(
            &mut input(measured()),
            &mut previous_layout,
            &previous.checkpoints,
            &[1, 1, 1],
            &[1, 1, 2],
            2,
        )
        .unwrap();
        assert_eq!(
            serde_json::to_string(&incremental.layout).unwrap(),
            serde_json::to_string(&previous.layout).unwrap()
        );
    }

    /// Lays `previous` out in full, then `next` incrementally with the blocks
    /// `dirty` changed, and checks it against a full pass over `next`.
    #[test]
    fn a_page_break_a_keep_with_next_follower_gains_relays_out_the_run_head() {
        let blocks = |page_break_before: bool| {
            vec![
                paragraph(0, 1, 90.0, json!({})),
                paragraph(1, 1, 10.0, json!({ "keepNext": true })),
                paragraph(2, 1, 10.0, json!({ "pageBreakBefore": page_break_before })),
            ]
        };
        let layout = assert_incremental_matches_full(blocks(false), blocks(true), &[2]);
        assert_eq!(layout.layout.pages.len(), 2);
        assert_incremental_matches_full(blocks(true), blocks(false), &[2]);

        let interior = |page_break_before: bool| {
            vec![
                paragraph(0, 1, 80.0, json!({})),
                paragraph(1, 1, 10.0, json!({ "keepNext": true })),
                paragraph(
                    2,
                    1,
                    10.0,
                    json!({ "keepNext": true, "pageBreakBefore": page_break_before }),
                ),
                paragraph(3, 1, 10.0, json!({})),
            ]
        };
        assert_incremental_matches_full(interior(false), interior(true), &[2]);
        assert_incremental_matches_full(interior(true), interior(false), &[2]);
    }

    fn assert_incremental_matches_full(
        previous: Vec<serde_json::Value>,
        next: Vec<serde_json::Value>,
        dirty: &[usize],
    ) -> CheckpointedLayout {
        assert_incremental_matches_full_with(json!({}), previous, next, dirty)
    }

    /// [`assert_incremental_matches_full`] with `options` over the test page's.
    fn assert_incremental_matches_full_with(
        options: serde_json::Value,
        previous: Vec<serde_json::Value>,
        next: Vec<serde_json::Value>,
        dirty: &[usize],
    ) -> CheckpointedLayout {
        let input = |measured: Vec<serde_json::Value>| {
            let mut value = input(measured);
            let mut merged = serde_json::to_value(&value.options).unwrap();
            for (key, option) in options.as_object().unwrap() {
                merged[key] = option.clone();
            }
            value.options = serde_json::from_value(merged).unwrap();
            value
        };
        let retained = layout_document_checkpointed(&mut input(previous)).unwrap();
        let previous_fingerprints = vec![1_u64; next.len()];
        let mut next_fingerprints = previous_fingerprints.clone();
        for &index in dirty {
            next_fingerprints[index] = 2;
        }
        let mut previous_layout = retained.layout.clone();
        let incremental = layout_document_incremental(
            &mut input(next.clone()),
            &mut previous_layout,
            &retained.checkpoints,
            &previous_fingerprints,
            &next_fingerprints,
            dirty[0],
        )
        .unwrap();
        let full = layout_document_checkpointed(&mut input(next)).unwrap();
        assert_eq!(
            serde_json::to_string(&incremental.layout).unwrap(),
            serde_json::to_string(&full.layout).unwrap()
        );
        assert_eq!(incremental.checkpoints, full.checkpoints);
        retained
    }

    #[test]
    fn placement_across_a_continuous_section_on_a_fresh_page_matches_a_full_pass() {
        let blocks = || {
            vec![
                paragraph(0, 1, 10.0, json!({})),
                paragraph(1, 1, 10.0, json!({ "pageBreakBefore": true })),
                json!({"block":{"kind":"pageBreak","id":"page"},"measure":{"kind":"pageBreak"}}),
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": "continuous",
                        "margins": { "top": 20, "right": 10, "bottom": 10, "left": 10 },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
                paragraph(2, 1, 10.0, json!({})),
                paragraph(3, 1, 10.0, json!({})),
            ]
        };
        assert_incremental_matches_full(blocks(), blocks(), &[0, 5]);
    }

    #[test]
    fn placement_does_not_resume_at_a_table_whose_row_moved_to_its_page() {
        let lines = 8;
        let table = json!({
            "block": {
                "kind": "table", "id": 90,
                "rows": [{ "id": 91, "cells": [{ "id": 92, "blocks": [paragraph(93, lines, 20.0, json!({ "keepLines": true }))["block"]] }] }],
                "columnWidths": [180],
            },
            "measure": {
                "kind": "table", "columnWidths": [180],
                "totalWidth": 180, "totalHeight": lines as f64 * 20.0,
                "rows": [{ "height": lines as f64 * 20.0, "cells": [{
                    "width": 180, "height": lines as f64 * 20.0,
                    "blocks": [paragraph(93, lines, 20.0, json!({ "keepLines": true }))["measure"]],
                }] }],
            },
        });
        let blocks = || {
            let mut blocks: Vec<_> = (0..24)
                .map(|id| paragraph(id, 1, 20.0, json!({})))
                .collect();
            blocks.push(paragraph(24, 1, 15.0, json!({})));
            blocks.push(table.clone());
            blocks.push(paragraph(25, 3, 20.0, json!({})));
            blocks
        };
        let full = layout_document_checkpointed(&mut input(blocks())).unwrap();
        assert!(
            full.checkpoints
                .iter()
                .any(|checkpoint| checkpoint.block_index == 25)
        );
        assert_incremental_matches_full(blocks(), blocks(), &[0, 26]);
        assert_incremental_matches_full(blocks(), blocks(), &[26]);
    }

    #[test]
    fn placement_restarts_from_the_origin_without_a_resumable_checkpoint() {
        let blocks = || vec![text_anchored_table(), paragraph(0, 1, 10.0, json!({}))];
        let retained = assert_incremental_matches_full(blocks(), blocks(), &[1]);
        assert_eq!(retained.layout.pages.len(), 2);
    }

    #[test]
    fn placement_does_not_resume_where_a_page_takes_new_section_geometry() {
        let blocks = || {
            vec![
                paragraph(0, 1, 10.0, json!({})),
                paragraph(1, 1, 10.0, json!({ "pageBreakBefore": true })),
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": "continuous",
                        "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
                paragraph(2, 9, 10.0, json!({})),
                paragraph(3, 1, 10.0, json!({ "alignment": "center" })),
                paragraph(4, 1, 10.0, json!({})),
            ]
        };
        let final_margins =
            json!({ "finalMargins": { "top": 10, "right": 20, "bottom": 10, "left": 20 } });
        let retained =
            assert_incremental_matches_full_with(final_margins, blocks(), blocks(), &[0, 5]);
        assert!(
            retained
                .layout
                .pages
                .windows(2)
                .any(|pair| pair[0].margins != pair[1].margins)
        );
    }

    #[test]
    fn a_changed_first_section_lays_out_from_the_origin() {
        let blocks = |top: f64| {
            vec![
                paragraph(0, 1, 10.0, json!({})),
                json!({
                    "block": {
                        "kind": "sectionBreak", "id": "section", "type": "nextPage",
                        "margins": { "top": top, "right": 10, "bottom": 10, "left": 10 },
                    },
                    "measure": { "kind": "sectionBreak" },
                }),
                paragraph(1, 1, 10.0, json!({})),
            ]
        };
        assert_incremental_matches_full(blocks(10.0), blocks(20.0), &[0, 1]);
    }

    #[test]
    fn incremental_placement_in_multi_column_sections_matches_a_full_pass() {
        let section = |columns: u32| {
            json!({
                "block": {
                    "kind": "sectionBreak", "id": format!("section{columns}"), "type": "nextPage",
                    "columns": { "count": columns, "gap": 20 },
                },
                "measure": { "kind": "sectionBreak" },
            })
        };
        for (blocks, dirty) in [
            (
                vec![
                    paragraph(0, 1, 10.0, json!({})),
                    section(1),
                    paragraph(1, 1, 10.0, json!({})),
                    paragraph(2, 1, 10.0, json!({})),
                    section(2),
                    paragraph(3, 1, 10.0, json!({})),
                ],
                &[3][..],
            ),
            (
                vec![
                    paragraph(0, 1, 10.0, json!({})),
                    section(2),
                    paragraph(1, 1, 10.0, json!({})),
                    paragraph(2, 1, 10.0, json!({})),
                ],
                &[3],
            ),
            (
                vec![
                    paragraph(0, 1, 10.0, json!({})),
                    paragraph(1, 1, 10.0, json!({ "pageBreakBefore": true })),
                    section(1),
                    paragraph(2, 1, 10.0, json!({})),
                    paragraph(3, 1, 10.0, json!({})),
                    section(2),
                    paragraph(4, 1, 10.0, json!({})),
                ],
                &[0, 6],
            ),
        ] {
            assert_incremental_matches_full(blocks.clone(), blocks, dirty);
        }
    }

    fn text_anchored_table() -> serde_json::Value {
        json!({
            "block": {
                "kind": "table", "id": 90,
                "rows": [{ "id": 91, "cantSplit": true, "cells": [{ "id": 92, "blocks": [] }] }],
                "columnWidths": [180],
                "floating": {
                    "vertAnchor": "text", "tblpY": 95,
                    "topFromText": 0, "bottomFromText": 0, "leftFromText": 0, "rightFromText": 0,
                },
            },
            "measure": {
                "kind": "table", "columnWidths": [180],
                "totalWidth": 180, "totalHeight": 20,
                "rows": [{ "height": 20, "cells": [
                    { "width": 180, "height": 20, "blocks": [] }
                ] }],
            },
        })
    }

    #[test]
    fn a_split_side_wrapped_float_and_its_anchor_match_incremental_layout() {
        let blocks = |height: f64, offset: f64, anchor_before: f64| {
            let floating = json!({
                "block": {
                    "kind": "table", "id": 90, "columnWidths": [60],
                    "rows": [
                        {"id": 91, "height": 30, "heightRule": "exact", "cells": []},
                        {"id": 92, "height": 30, "heightRule": "exact", "cells": []}
                    ],
                    "floating": {
                        "horzAnchor": "text", "vertAnchor": "text",
                        "tblpXSpec": "right", "tblpY": offset,
                        "leftFromText": 10, "rightFromText": 10
                    }
                },
                "measure": {
                    "kind": "table", "columnWidths": [60], "totalWidth": 60, "totalHeight": 60,
                    "rows": [{"height": 30, "cells": []}, {"height": 30, "cells": []}]
                }
            });
            let mut anchor = paragraph(2, 3, 10.0, json!({"spacing": {"before": anchor_before}}));
            for line in anchor["measure"]["lines"].as_array_mut().unwrap() {
                line["rightOffset"] = json!(70);
            }
            vec![
                paragraph(0, 1, 10.0, json!({})),
                paragraph(
                    1,
                    1,
                    height,
                    json!({"pageBreakBefore": true, "spacing": {"after": 12}}),
                ),
                floating,
                anchor,
            ]
        };
        let split = blocks(40.0, 10.0, 0.0);
        let spaced = blocks(40.0, 10.0, 12.0);
        assert_incremental_matches_full(split.clone(), spaced.clone(), &[3]);
        let retained = assert_incremental_matches_full(spaced, split.clone(), &[3]);
        assert_eq!(retained.layout.pages.len(), 3);
        let Some(Fragment::Table(first)) = retained.layout.pages[1].fragments.last() else {
            panic!("first table fragment expected");
        };
        assert_eq!((first.y, first.row_start, first.row_end), (60.0, 0, 1));
        assert_eq!(first.carried_to_next, Some(true));
        let [Fragment::Table(table), Fragment::Paragraph(anchor)] =
            retained.layout.pages[2].fragments.as_slice()
        else {
            panic!("table and anchor expected");
        };
        assert_eq!(
            (table.x, table.y, table.row_start, table.row_end),
            (130.0, 10.0, 1, 2)
        );
        assert_eq!((anchor.y, anchor.from_line), (10.0, 0));
        let fitting = blocks(10.0, 10.0, 0.0);
        assert_incremental_matches_full(split.clone(), fitting.clone(), &[1]);
        assert_incremental_matches_full(fitting, split.clone(), &[1]);
        assert_incremental_matches_full(split, blocks(40.0, 45.0, 0.0), &[2]);
    }

    #[test]
    fn placement_does_not_resume_at_a_floating_table_that_opened_its_page() {
        let blocks = |page_break: bool| {
            vec![
                paragraph(0, 1, 10.0, json!({})),
                paragraph(1, 1, 10.0, json!({ "pageBreakBefore": true })),
                if page_break {
                    json!({"block":{"kind":"pageBreak","id":"page"},"measure":{"kind":"pageBreak"}})
                } else {
                    json!({"block":{"kind":"columnBreak","id":"column"},"measure":{"kind":"columnBreak"}})
                },
                text_anchored_table(),
                paragraph(2, 1, 10.0, json!({})),
            ]
        };
        let retained = assert_incremental_matches_full(blocks(true), blocks(true), &[0, 4]);
        assert_eq!(retained.layout.pages.len(), 4);
        assert_incremental_matches_full(blocks(true), blocks(false), &[2]);
    }

    #[test]
    fn incremental_checkpoints_match_a_full_pass_on_a_page_with_several() {
        let page_break =
            || json!({"block":{"kind":"pageBreak","id":"page"},"measure":{"kind":"pageBreak"}});
        let measured = || {
            let mut blocks = vec![paragraph(0, 1, 10.0, json!({})), page_break(), page_break()];
            blocks.extend((1..40).map(|id| paragraph(id, 1, 10.0, json!({}))));
            blocks
        };
        let previous = layout_document_checkpointed(&mut input(measured())).unwrap();
        // The second break opens no page: its checkpoint and the next paragraph's share one.
        assert!(
            previous
                .checkpoints
                .windows(2)
                .any(|pair| pair[0].page_index == pair[1].page_index)
        );
        let previous_fingerprints = vec![1_u64; measured().len()];
        let mut next_fingerprints = previous_fingerprints.clone();
        next_fingerprints[2] = 2;
        next_fingerprints[35] = 2;
        let mut previous_layout = previous.layout.clone();
        let incremental = layout_document_incremental(
            &mut input(measured()),
            &mut previous_layout,
            &previous.checkpoints,
            &previous_fingerprints,
            &next_fingerprints,
            2,
        )
        .unwrap();
        assert_eq!(incremental.checkpoints, previous.checkpoints);
        assert_eq!(
            serde_json::to_string(&incremental.layout).unwrap(),
            serde_json::to_string(&previous.layout).unwrap()
        );
    }

    #[test]
    fn a_heading_above_a_table_on_a_page_with_float_bands_keeps_the_whole_row_witness() {
        let cell_paragraph = json!({
            "kind": "paragraph", "id": 10,
            "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
        });
        let lines: Vec<_> = [25.0, 25.0, 20.0, 20.0].into_iter().map(line).collect();
        let table = json!({
            "block": {
                "kind": "table", "id": 3,
                "rows": [{ "id": 20, "cells": [{ "id": 30, "blocks": [cell_paragraph] }] }],
                "columnWidths": [100],
            },
            "measure": {
                "kind": "table", "columnWidths": [100], "totalWidth": 100, "totalHeight": 90,
                "rows": [{ "height": 90, "cells": [{ "width": 100, "height": 90, "blocks": [
                    { "kind": "paragraph", "lines": lines, "totalHeight": 90 }
                ] }] }],
            },
        });
        let mut value = input(vec![
            paragraph(1, 1, 5.0, json!({})),
            paragraph(2, 1, 15.0, json!({ "keepNext": true })),
            table,
        ]);
        value.options.section_page_float_bands = Some(
            serde_json::from_value(json!([{"default": [{"top": 60, "bottom": 65}]}])).unwrap(),
        );
        let result = layout_document(&mut value).unwrap();
        assert_eq!(result.pages.len(), 2);
        assert!(result.pages[0].fragments.iter().any(|fragment| matches!(
            fragment,
            Fragment::Paragraph(p)
                if matches!(p.block_id, crate::types::BlockId::Num(value) if value == 2.0)
        )));
    }

    fn oversized_cant_split_table() -> serde_json::Value {
        let paragraph_block = json!({
            "kind": "paragraph", "id": 10,
            "runs": [{ "kind": "text", "text": "x", "fmt": {} }],
        });
        let paragraph_extent = json!({
            "kind": "paragraph",
            "lines": vec![line(20.0); 10],
            "totalHeight": 200,
        });
        json!({
            "block": {
                "kind": "table", "id": 2,
                "rows": [{ "id": 20, "cantSplit": true, "cells": [
                    { "id": 30, "blocks": [paragraph_block] }
                ] }],
                "columnWidths": [100],
            },
            "measure": {
                "kind": "table", "columnWidths": [100],
                "totalWidth": 100, "totalHeight": 200,
                "rows": [{ "height": 200, "cells": [
                    { "width": 100, "height": 200, "blocks": [paragraph_extent] }
                ] }],
            },
        })
    }

    fn positioned_floating_table() -> serde_json::Value {
        json!({
            "block": {
                "kind": "table", "id": 3,
                "rows": [{ "id": 1, "cells": [{ "id": 2, "blocks": [] }] }],
                "columnWidths": [50],
                "floating": {
                    "horzAnchor": "page", "tblpXSpec": "right",
                    "vertAnchor": "page", "tblpY": 5,
                },
            },
            "measure": {
                "kind": "table", "columnWidths": [50],
                "totalWidth": 50, "totalHeight": 40,
                "rows": [{ "height": 40, "cells": [
                    { "width": 50, "height": 40, "blocks": [] }
                ] }],
            },
        })
    }

    fn positioned_image() -> serde_json::Value {
        json!({
            "block": {
                "kind": "image", "id": 4, "src": "embedded", "width": 50, "height": 20,
                "anchor": {
                    "isAnchored": true,
                    "position": {
                        "horizontal": { "relativeTo": "page", "align": "center" },
                        "vertical": { "relativeTo": "page", "posOffset": 5 }
                    }
                }
            },
            "measure": { "kind": "image", "width": 50, "height": 20 }
        })
    }

    #[test]
    fn oversized_keep_lines_terminates_and_remains_visible() {
        let result = layout(vec![paragraph(1, 10, 40.0, json!({ "keepLines": true }))]);
        assert!(!result.pages.is_empty());
        let fragments = result
            .pages
            .iter()
            .flat_map(|page| page.fragments.iter())
            .count();
        assert_eq!(fragments, 5);
    }

    #[test]
    fn widow_control_advances_a_single_bottom_line_and_keeps_two_at_each_side() {
        let result = layout(vec![
            paragraph(1, 1, 70.0, json!({})),
            paragraph(2, 4, 20.0, json!({})),
        ]);
        assert_eq!(result.pages.len(), 2);
        let second_page_lines: Vec<(usize, usize)> = result.pages[1]
            .fragments
            .iter()
            .filter_map(|fragment| match fragment {
                Fragment::Paragraph(p)
                    if matches!(p.block_id, crate::types::BlockId::Num(value) if value == 2.0) =>
                {
                    Some((p.from_line, p.to_line))
                }
                _ => None,
            })
            .collect();
        assert_eq!(second_page_lines, vec![(0, 4)]);
    }

    #[test]
    fn widow_control_carries_the_pushed_line_to_the_next_page() {
        for (preceding_height, lines, expected) in [
            (40.0, 4, vec![(0, 0, 2), (1, 2, 4)]),
            (20.0, 5, vec![(0, 0, 3), (1, 3, 5)]),
        ] {
            let result = layout(vec![
                paragraph(1, 1, preceding_height, json!({})),
                paragraph(2, lines, 20.0, json!({})),
            ]);
            assert_eq!(paragraph_slices(&result, 2.0), expected);
        }
    }

    #[test]
    fn widow_control_keeps_the_pushed_line_when_the_last_two_cannot_share_a_page() {
        let mut block = paragraph(1, 4, 20.0, json!({}));
        block["measure"]["lines"][3]["lineHeight"] = json!(90.0);
        let result = layout(vec![block]);
        assert_eq!(result.pages.len(), 2);
    }

    #[test]
    fn widow_control_keeps_the_pushed_line_above_a_float_band_on_its_page() {
        let mut value = input(vec![paragraph(1, 4, 20.0, json!({}))]);
        value.options.page_size = Some(crate::types::Size { w: 200.0, h: 220.0 });
        value.options.section_page_float_bands = Some(
            serde_json::from_value(json!([{"default": [{"top": 70, "bottom": 90}]}])).unwrap(),
        );
        let result = layout_document(&mut value).unwrap();
        assert_eq!(result.pages.len(), 1);
    }

    #[test]
    fn widow_control_in_balanced_columns_keeps_the_section_on_one_page() {
        let mut value = input(vec![
            paragraph(1, 4, 20.0, json!({})),
            paragraph(2, 5, 20.0, json!({})),
            paragraph(3, 4, 20.0, json!({})),
            paragraph(4, 5, 20.0, json!({})),
        ]);
        value.options.page_size = Some(crate::types::Size { w: 500.0, h: 320.0 });
        value.options.columns =
            Some(serde_json::from_value(json!({"count": 3, "gap": 20})).unwrap());
        let result = layout_document(&mut value).unwrap();
        assert_eq!(result.pages.len(), 1);
    }

    fn paragraph_slices(layout: &Layout, id: f64) -> Vec<(usize, usize, usize)> {
        layout
            .pages
            .iter()
            .enumerate()
            .flat_map(|(page_index, page)| {
                page.fragments.iter().filter_map(move |fragment| match fragment {
                    Fragment::Paragraph(p)
                        if matches!(p.block_id, crate::types::BlockId::Num(value) if value == id) =>
                    {
                        Some((page_index, p.from_line, p.to_line))
                    }
                    _ => None,
                })
            })
            .collect()
    }

    #[test]
    fn widow_control_keeps_short_paragraphs_together() {
        for (lines, preceding_height) in [(2, 70.0), (3, 50.0), (3, 70.0)] {
            let result = layout(vec![
                paragraph(1, 1, preceding_height, json!({})),
                paragraph(2, lines, 20.0, json!({})),
            ]);
            assert_eq!(paragraph_slices(&result, 2.0), vec![(1, 0, lines)]);
        }
    }

    #[test]
    fn short_paragraphs_still_split_when_widow_control_is_disabled() {
        for (lines, preceding_height, split) in [(2, 70.0, 1), (3, 50.0, 2), (3, 70.0, 1)] {
            let result = layout(vec![
                paragraph(1, 1, preceding_height, json!({})),
                paragraph(2, lines, 20.0, json!({ "widowControl": false })),
            ]);
            assert_eq!(
                paragraph_slices(&result, 2.0),
                vec![(0, 0, split), (1, split, lines)]
            );
        }
    }

    #[test]
    fn widow_control_preserves_short_paragraphs_that_fit_exactly() {
        for lines in [2, 3] {
            let result = layout(vec![
                paragraph(1, 1, 100.0 - lines as f64 * 20.0, json!({})),
                paragraph(2, lines, 20.0, json!({})),
            ]);
            assert_eq!(paragraph_slices(&result, 2.0), vec![(0, 0, lines)]);
        }
    }

    #[test]
    fn widow_control_discards_boundary_spacing_when_moving_short_paragraphs() {
        let result = layout(vec![
            paragraph(1, 1, 30.0, json!({ "spacing": { "after": 60 } })),
            paragraph(2, 3, 20.0, json!({ "spacing": { "before": 50 } })),
        ]);
        assert_eq!(paragraph_slices(&result, 2.0), vec![(1, 0, 3)]);
        let Fragment::Paragraph(fragment) = &result.pages[1].fragments[0] else {
            panic!()
        };
        assert_eq!(fragment.y, 10.0);
    }

    #[test]
    fn oversized_short_paragraphs_still_terminate_with_every_line_visible() {
        for (lines, height, expected) in [
            (2, 60.0, vec![(0, 0, 1), (1, 1, 2)]),
            (3, 60.0, vec![(0, 0, 1), (1, 1, 2), (2, 2, 3)]),
            (3, 40.0, vec![(0, 0, 2), (1, 2, 3)]),
        ] {
            let result = layout(vec![paragraph(1, lines, height, json!({}))]);
            assert_eq!(paragraph_slices(&result, 1.0), expected);
        }
    }

    #[test]
    fn authored_widow_control_off_splits_where_the_default_moves_the_paragraph_on() {
        let default = layout(vec![
            paragraph(1, 1, 70.0, json!({})),
            paragraph(2, 4, 20.0, json!({})),
        ]);
        let disabled = layout(vec![
            paragraph(1, 1, 70.0, json!({})),
            paragraph(2, 4, 20.0, json!({ "widowControl": false })),
        ]);

        assert_eq!(paragraph_slices(&default, 2.0), vec![(1, 0, 4)]);
        assert_eq!(paragraph_slices(&disabled, 2.0), vec![(0, 0, 1), (1, 1, 4)]);
    }

    #[test]
    fn authored_widow_control_off_saves_the_page_the_default_costs() {
        let default = layout(vec![
            paragraph(1, 1, 70.0, json!({})),
            paragraph(2, 4, 20.0, json!({})),
            paragraph(3, 2, 20.0, json!({})),
        ]);
        let disabled = layout(vec![
            paragraph(1, 1, 70.0, json!({})),
            paragraph(2, 4, 20.0, json!({ "widowControl": false })),
            paragraph(3, 2, 20.0, json!({})),
        ]);

        assert_eq!(default.pages.len(), 3);
        assert_eq!(paragraph_slices(&default, 3.0), vec![(2, 0, 2)]);
        assert_eq!(disabled.pages.len(), 2);
        assert_eq!(paragraph_slices(&disabled, 3.0), vec![(1, 0, 2)]);
    }

    #[test]
    fn unavoidable_oversized_cant_split_row_terminates_as_visible_safe_slices() {
        let result = layout(vec![oversized_cant_split_table()]);
        assert_eq!(result.pages.len(), 2);
        let fragments: Vec<&crate::types::TableFragment> = result
            .pages
            .iter()
            .flat_map(|page| page.fragments.iter())
            .filter_map(|fragment| match fragment {
                Fragment::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(fragments.len(), 2);
        assert_eq!(fragments[0].clip_bottom, Some(100.0));
        assert_eq!(fragments[1].clip_top, Some(100.0));
    }

    #[test]
    fn floating_tables_and_anchored_images_share_page_relative_placement_semantics() {
        let result = layout(vec![positioned_floating_table(), positioned_image()]);
        let page = &result.pages[0];
        let table = page
            .fragments
            .iter()
            .find_map(|fragment| match fragment {
                Fragment::Table(value) => Some(value),
                _ => None,
            })
            .unwrap();
        assert_eq!(
            (table.x, table.y, table.is_floating),
            (150.0, 5.0, Some(true))
        );
        let image = page
            .fragments
            .iter()
            .find_map(|fragment| match fragment {
                Fragment::Image(value) => Some(value),
                _ => None,
            })
            .unwrap();
        assert_eq!((image.x, image.y), (75.0, 5.0));
    }
}
