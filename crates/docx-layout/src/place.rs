//! The placement walk — port of `layoutDocument`/`place` and the per-kind
//! placers in `packages/core/src/layout/pagination/index.ts`.
//!
//! One pass over the measured blocks, a per-kind placer for each, the
//! paginator holding page/column state, and all look-ahead answered from the
//! prescan plan. Paragraph placement (including spacing collapse, float line
//! offsets, page/column splits, per-fragment PM ranges and resolved lines) is
//! fully ported; the features listed in `hooks.rs` are stubbed there and make
//! the engine return `Unsupported` so the JS seam falls back to TypeScript.
//!
//! Checkpoint capture (`controls.checkpoints` in the TS walk) is not ported:
//! checkpoints are derived resume bookmarks, excluded from golden
//! serialization, and only consumed by incremental pagination which stays on
//! the TS side for now. They influence no fragment or page output.
//!
//! Also ported here, because the spine depends on them:
//! - `getParagraphFragmentPmRange` from `paragraphFragmentRange.ts`,
//! - `isFloatingTextBoxBlock` from `textBoxFlow.ts` (+ `docx/wrapTypes.ts`).
//!
//! `resolvePageMargins` lives in `section_breaks.rs` and the spacing helpers
//! in `paragraph_spacing.rs`, mirroring the TS module layout.

use crate::LayoutError;
use crate::hooks;
use crate::page_flow::{PageFlowGeometry, Paginator};
use crate::paragraph_spacing::{get_spacing_after, get_spacing_before};
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

struct ConvergenceInput<'a> {
    previous_checkpoints: &'a [LayoutCheckpoint],
    previous_fingerprints: &'a [u64],
    next_fingerprints: &'a [u64],
    dirty_index: usize,
}

impl ConvergenceInput<'_> {
    fn retained_match(&self, checkpoint: &LayoutCheckpoint) -> Option<&LayoutCheckpoint> {
        if checkpoint.block_index <= self.dirty_index
            || self.previous_fingerprints.get(checkpoint.block_index..)
                != self.next_fingerprints.get(checkpoint.block_index..)
        {
            return None;
        }
        self.previous_checkpoints.iter().find(|previous| {
            previous.block_index == checkpoint.block_index
                && previous.section_index == checkpoint.section_index
                && previous.page_number == checkpoint.page_number
                && previous.flow == checkpoint.flow
        })
    }
}

// ---------------------------------------------------------------------------
// spine helpers ported from sibling TS modules
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq)]
enum Edge {
    Start,
    End,
}

/// `runBoundaryPmPos` (paragraphFragmentRange.ts).
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

/// `getParagraphFragmentPmRange` (paragraphFragmentRange.ts).
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

/// `isFloatingWrapType` (docx/wrapTypes.ts).
fn is_floating_wrap_type(wrap_type: Option<&str>) -> bool {
    matches!(
        wrap_type,
        Some("square") | Some("tight") | Some("through") | Some("behind") | Some("inFront")
    )
}

/// `isFloatingTextBoxBlock` (textBoxFlow.ts).
fn is_floating_text_box_block(block: &TextBoxBlock) -> bool {
    block.display_mode.as_deref() == Some("float")
        || is_floating_wrap_type(block.wrap_type.as_deref())
        || block.wrap_type.as_deref() == Some("topAndBottom")
}

// ---------------------------------------------------------------------------
// contextual spacing (index.ts applyContextualSpacing)
// ---------------------------------------------------------------------------

/// Suppress spacing between two adjacent same-style contextual paragraphs
/// (OOXML §17.3.1.9). Mutates in place, exactly like the TS.
fn contextual_spacing_pair(curr: &mut LayoutBlock, next: &mut LayoutBlock) {
    let (LayoutBlock::Paragraph(c), LayoutBlock::Paragraph(n)) = (curr, next) else {
        return;
    };
    let same_style = c
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.style_id.as_deref())
        .unwrap_or("")
        == n.attrs
            .as_ref()
            .and_then(|attrs| attrs.style_id.as_deref())
            .unwrap_or("");
    if !same_style {
        return;
    }
    if let Some(ca) = &mut c.attrs
        && ca.contextual_spacing.unwrap_or(false)
        && let Some(spacing) = &mut ca.spacing
    {
        spacing.after = Some(0.0);
    }
    if let Some(na) = &mut n.attrs
        && na.contextual_spacing.unwrap_or(false)
        && let Some(spacing) = &mut na.spacing
    {
        spacing.before = Some(0.0);
    }
}

/// `applyContextualSpacing` over a plain block list (table-cell recursion).
fn apply_contextual_spacing_blocks(blocks: &mut [LayoutBlock]) {
    for i in 0..blocks.len().saturating_sub(1) {
        let (head, tail) = blocks.split_at_mut(i + 1);
        contextual_spacing_pair(&mut head[i], &mut tail[0]);
    }
    for block in blocks.iter_mut() {
        if let LayoutBlock::Table(table) = block {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    apply_contextual_spacing_blocks(&mut cell.blocks);
                }
            }
        }
    }
}

/// `applyContextualSpacing` over the measured list (TS mutates the shared
/// block references; here the blocks live inside `MeasuredBlock`).
fn apply_contextual_spacing_measured(measured: &mut [MeasuredBlock]) {
    for i in 0..measured.len().saturating_sub(1) {
        let (head, tail) = measured.split_at_mut(i + 1);
        contextual_spacing_pair(&mut head[i].block, &mut tail[0].block);
    }
    for mb in measured.iter_mut() {
        if let LayoutBlock::Table(table) = &mut mb.block {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    apply_contextual_spacing_blocks(&mut cell.blocks);
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// entry point (index.ts layoutDocument)
// ---------------------------------------------------------------------------

/// `layoutDocument` — convert measured blocks into pages with positioned
/// fragments, or `Unsupported` when a not-yet-ported feature is engaged.
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

    let plan = prescan(
        measured,
        &body_config,
        final_config,
        options.body_break_type,
    )?;

    let initial_config = plan.section_configs.first().cloned().unwrap_or(body_config);

    let mut paginator = Paginator::new(
        initial_config.page_size.clone(),
        initial_config.margins.clone(),
        initial_config
            .columns
            .clone()
            .unwrap_or_else(default_columns),
        options.footnote_reserved_heights.clone(),
    )?;

    let placement = place(
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
pub fn layout_document_incremental(
    input: &mut Input,
    previous_layout: &Layout,
    previous_checkpoints: &[LayoutCheckpoint],
    previous_fingerprints: &[u64],
    next_fingerprints: &[u64],
    dirty_index: usize,
) -> Result<CheckpointedLayout, LayoutError> {
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
    let plan = prescan(
        &input.measured,
        &body_config,
        final_config,
        options.body_break_type,
    )?;
    let initial_config = plan.section_configs.first().cloned().unwrap_or(body_config);
    let resume = previous_checkpoints
        .iter()
        .rev()
        .find(|checkpoint| checkpoint.block_index <= dirty_index)
        .ok_or_else(|| LayoutError::Unsupported("no clean pagination checkpoint".into()))?;
    let prefix_pages = previous_layout.pages[..resume.page_index].to_vec();
    let prefix_checkpoints: Vec<_> = previous_checkpoints
        .iter()
        .filter(|checkpoint| checkpoint.page_index < resume.page_index)
        .cloned()
        .collect();
    let mut paginator = Paginator::resume(
        &resume.flow,
        resume.page_number,
        options.footnote_reserved_heights.clone(),
    )?;
    paginator.set_section_index(resume.section_index);
    let convergence = ConvergenceInput {
        previous_checkpoints,
        previous_fingerprints,
        next_fingerprints,
        dirty_index,
    };
    let placement = place(
        &input.measured,
        &plan,
        &mut paginator,
        &initial_config,
        resume.block_index,
        resume.section_index,
        resume.page_index,
        Some(&convergence),
    )?;

    let rebuilt_page_end = placement
        .converged
        .as_ref()
        .map_or(resume.page_index + paginator.pages.len(), |(next, _)| {
            next.page_index
        });
    let mut pages = prefix_pages;
    pages.append(&mut paginator.pages);
    let mut checkpoints = prefix_checkpoints;
    checkpoints.extend(placement.checkpoints);

    if let Some((next_checkpoint, previous_checkpoint)) = placement.converged {
        debug_assert_eq!(pages.len(), next_checkpoint.page_index);
        let reused_page_start = pages.len();
        pages.extend_from_slice(&previous_layout.pages[previous_checkpoint.page_index..]);
        refresh_reused_paragraph_pages(&mut pages[reused_page_start..], &input.measured);
        let page_shift =
            next_checkpoint.page_index as isize - previous_checkpoint.page_index as isize;
        checkpoints.extend(
            previous_checkpoints
                .iter()
                .filter(|checkpoint| checkpoint.page_index >= previous_checkpoint.page_index)
                .cloned()
                .map(|mut checkpoint| {
                    checkpoint.page_index = (checkpoint.page_index as isize + page_shift) as usize;
                    checkpoint
                }),
        );
    }

    Ok(CheckpointedLayout {
        layout: Layout {
            page_size,
            pages,
            columns: options.columns.clone(),
            headers: None,
            footers: None,
            page_gap: options.page_gap,
        },
        checkpoints,
        placed_blocks: placement.placed_blocks,
        rebuilt_page_start: resume.page_index,
        rebuilt_page_end,
    })
}

// ---------------------------------------------------------------------------
// placement walk (index.ts place)
// ---------------------------------------------------------------------------

struct PlacementOutcome {
    checkpoints: Vec<LayoutCheckpoint>,
    placed_blocks: usize,
    converged: Option<(LayoutCheckpoint, LayoutCheckpoint)>,
}

fn place(
    measured: &[MeasuredBlock],
    plan: &LayoutPlan,
    paginator: &mut Paginator,
    initial_config: &SectionLayoutConfig,
    start_index: usize,
    mut section_idx: usize,
    page_index_offset: usize,
    convergence: Option<&ConvergenceInput<'_>>,
) -> Result<PlacementOutcome, LayoutError> {
    let mut checkpoints = Vec::new();
    let mut placed_blocks = 0usize;

    if initial_config
        .columns
        .as_ref()
        .map_or(1.0, |columns| columns.count)
        > 1.0
    {
        hooks::balance_terminal_continuous_text_columns(
            measured,
            paginator,
            start_index,
            plan.break_indices
                .first()
                .copied()
                .unwrap_or(measured.len()),
        )?;
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
            if let Some(previous) = convergence.and_then(|value| value.retained_match(&checkpoint))
            {
                paginator.pages.truncate(page_index);
                return Ok(PlacementOutcome {
                    checkpoints,
                    placed_blocks,
                    converged: Some((checkpoint, previous.clone())),
                });
            }
            checkpointed_page = Some(page_index);
            checkpoints.push(checkpoint);
        }
        let fragments_before = paginator.page_fragment_counts();
        // pageBreakBefore forces a fresh page before the block is placed
        if hooks::breaks_before_block(&mb.block)? {
            paginator.force_page_break();
        }

        // at the head of a keep-with-next group, move to a fresh page when the
        // whole group would otherwise straddle the boundary
        if let Some(group) = plan.keep_with_next.groups_by_head.get(&i)
            && !plan.keep_with_next.interior_members.contains(&i)
        {
            let state_idx = paginator.get_current();
            let page_content_height =
                paginator.state(state_idx).content_limit - paginator.state(state_idx).content_top;
            let page_has_content = paginator.page_fragment_count(state_idx) > 0;
            let group_height = hooks::measure_keep_with_next_group(group, measured)?;
            let must_advance = hooks::keep_with_next_group_must_advance(
                group_height,
                paginator.get_available_height(),
                page_content_height,
                page_has_content,
            )?;
            if must_advance {
                paginator.force_page_break();
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
                paginator.force_page_break();
            }

            LayoutBlock::ColumnBreak(_) => {
                paginator.force_column_break();
            }

            LayoutBlock::SectionBreak(block) => {
                // use the NEXT section's columns; for break type, prefer the
                // next section's but fall back to the current break's
                let next_type: Option<SectionBreakType> = plan
                    .section_break_types
                    .get(section_idx + 1)
                    .copied()
                    .flatten()
                    .or_else(|| plan.section_break_types.get(section_idx).copied().flatten());
                let next_section_config = plan
                    .section_configs
                    .get(section_idx + 1)
                    .cloned()
                    .unwrap_or_else(|| initial_config.clone());
                paginator.set_section_index(section_idx + 1);
                hooks::handle_section_break(block, paginator, &next_section_config, next_type)?;

                let next_break_index = plan.break_indices.get(section_idx + 1).copied();
                if next_section_config
                    .columns
                    .as_ref()
                    .map_or(1.0, |c| c.count)
                    > 1.0
                {
                    hooks::balance_terminal_continuous_text_columns(
                        measured,
                        paginator,
                        i + 1,
                        next_break_index.unwrap_or(measured.len()),
                    )?;
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
            if let Some(previous) = convergence.and_then(|value| value.retained_match(&checkpoint))
            {
                paginator.pages.truncate(page_index);
                return Ok(PlacementOutcome {
                    checkpoints,
                    placed_blocks,
                    converged: Some((checkpoint, previous.clone())),
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

/// Retained suffix pages keep their geometry but absolute PM positions move
/// after an earlier edit. Refresh paragraph fragment ranges and resolved run
/// slices from the new measured arena before the display list consumes them.
fn refresh_reused_paragraph_pages(pages: &mut [crate::types::Page], measured: &[MeasuredBlock]) {
    let paragraphs: std::collections::HashMap<_, _> = measured
        .iter()
        .filter_map(|measured| match (&measured.block, &measured.measure) {
            (LayoutBlock::Paragraph(block), BlockExtent::Paragraph(extent)) => {
                Some((block_id_key(&block.id), (block, extent)))
            }
            _ => None,
        })
        .collect();
    for page in pages {
        for fragment in &mut page.fragments {
            let Fragment::Paragraph(fragment) = fragment else {
                continue;
            };
            let Some((block, extent)) = paragraphs.get(&block_id_key(&fragment.block_id)) else {
                continue;
            };
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
    }
}

// ---------------------------------------------------------------------------
// per-kind placers
// ---------------------------------------------------------------------------

/// `buildResolvedLines` — materialize resolved run segments for the
/// fragment's lines `[from_line, to_line)`.
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

/// `layoutParagraph` — place a paragraph's measured lines, splitting into
/// carried fragments whenever the page or column runs out of room.
fn layout_paragraph(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    paginator: &mut Paginator,
) -> Result<(), LayoutError> {
    // a run kind this port doesn't know can't be re-emitted faithfully in
    // resolved lines — degrade to the TS engine
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
            y: pen_y + space_before,
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

    if block
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.keep_lines)
        .unwrap_or(false)
    {
        let state_idx = paginator.get_current();
        let state = paginator.state(state_idx);
        let capacity = state.content_limit - state.content_top;
        let required = space_before.max(state.deferred_spacing) + paragraph_height;
        if required <= capacity && required > paginator.get_available_height() {
            paginator.ensure_fits(required);
        }
    }

    let mut current_line_index = 0usize;

    while current_line_index < lines.len() {
        let state_idx = paginator.get_current();
        let deferred_spacing = paginator.state(state_idx).deferred_spacing;
        let column_index = paginator.state(state_idx).column_index;

        // reserve the space addFragment will consume before this fragment's
        // first line so the greedy fit budgets against what actually remains
        let reserved_before = if current_line_index == 0 {
            space_before.max(deferred_spacing)
        } else {
            0.0
        };
        let available_for_lines = paginator.get_available_height() - reserved_before;

        // greedy fit; a fragment always takes at least one line
        let mut lines_height = 0.0f64;
        let mut fitting_lines = 0usize;

        for line in &lines[current_line_index..] {
            // floatSkipBefore counts toward fragment height so following
            // blocks flow below the float, not over it
            let line_height = line.line_height + line.float_skip_before.unwrap_or(0.0);
            let total_with_line = lines_height + line_height;

            if total_with_line <= available_for_lines || fitting_lines == 0 {
                lines_height = total_with_line;
                fitting_lines += 1;
            } else {
                break;
            }
        }

        let remaining_after = lines.len() - (current_line_index + fitting_lines);
        let widow_control = lines.len() >= 4;
        if widow_control && remaining_after > 0 {
            if current_line_index == 0 && fitting_lines == 1 {
                let capacity = paginator.state(state_idx).content_limit
                    - paginator.state(state_idx).content_top;
                let first_two_height = lines.iter().take(2).fold(0.0, |sum, line| {
                    sum + line.line_height + line.float_skip_before.unwrap_or(0.0)
                });
                if reserved_before + first_two_height <= capacity {
                    paginator.force_column_break();
                    continue;
                }
            }
            if remaining_after == 1 && fitting_lines > 2 {
                fitting_lines -= 1;
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

        // leftover lines: move the pen to a column/page with room for the next
        if current_line_index < lines.len() {
            paginator.ensure_fits(lines[current_line_index].line_height);
        }
    }

    Ok(())
}

/// `layoutImage` — inline images consume flow height; anchored ones route to
/// the anchored placer and float over the page instead.
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

/// Common DrawingML shapes consume their measured bbox in normal flow. The
/// bridge currently emits the PM path's in-flow shape contract; exotic anchor
/// scenes remain eligible for the PM fallback before they reach pagination.
fn layout_shape(block: &ShapeBlock, measure: &ShapeExtent, paginator: &mut Paginator) {
    let state_idx = paginator.ensure_fits(measure.height);
    let column_index = paginator.state(state_idx).column_index;
    let fragment = Fragment::Shape(ShapeFragment {
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

fn resolve_object_position(
    position: Option<&ImageRunPosition>,
    width: f64,
    height: f64,
    paginator: &mut Paginator,
) -> (f64, f64) {
    let state_idx = paginator.get_current();
    let state = paginator.state(state_idx);
    let page = &paginator.pages[state.page_index];
    let column_x = paginator.get_column_x(state.column_index);
    if let Some(position) = position
        && position.use_simple_pos.unwrap_or(false)
        && let Some(simple) = position
            .simple_pos
            .as_ref()
            .and_then(|value| value.as_object())
    {
        let x = simple
            .get("x")
            .and_then(|value| value.as_f64())
            .filter(|value| value.is_finite())
            .unwrap_or(column_x);
        let y = simple
            .get("y")
            .and_then(|value| value.as_f64())
            .filter(|value| value.is_finite())
            .unwrap_or(state.pen_y);
        return (x, y);
    }

    let coordinate = |spec: Option<&crate::types::AxisPosition>, horizontal: bool| {
        let relative_to = spec
            .and_then(|axis| axis.relative_to.as_deref())
            .unwrap_or(if horizontal { "column" } else { "paragraph" });
        let odd = page.number % 2 == 1;
        let (start, end) = if horizontal {
            match relative_to {
                "page" => (0.0, page.size.w),
                "margin" => (page.margins.left, page.size.w - page.margins.right),
                "leftMargin" => (0.0, page.margins.left),
                "rightMargin" => (page.size.w - page.margins.right, page.size.w),
                "insideMargin" if odd => (0.0, page.margins.left),
                "insideMargin" => (page.size.w - page.margins.right, page.size.w),
                "outsideMargin" if odd => (page.size.w - page.margins.right, page.size.w),
                "outsideMargin" => (0.0, page.margins.left),
                _ => (column_x, column_x + paginator.column_width()),
            }
        } else {
            match relative_to {
                "page" => (0.0, page.size.h),
                "margin" => (page.margins.top, page.size.h - page.margins.bottom),
                "topMargin" => (0.0, page.margins.top),
                "bottomMargin" => (page.size.h - page.margins.bottom, page.size.h),
                _ => (state.pen_y, state.content_limit),
            }
        };
        let extent = if horizontal { width } else { height };
        if let Some(offset) = spec
            .and_then(|axis| axis.pos_offset)
            .filter(|value| value.is_finite())
        {
            return start + offset;
        }
        match spec.and_then(|axis| axis.align.as_deref()) {
            Some("center") => start + (end - start - extent) / 2.0,
            Some("right" | "bottom" | "outside") => end - extent,
            Some("inside") if !odd => end - extent,
            _ => start,
        }
    };
    (
        coordinate(position.and_then(|value| value.horizontal.as_ref()), true),
        coordinate(position.and_then(|value| value.vertical.as_ref()), false),
    )
}

/// `layoutAnchoredImage` — absolute placement at the anchor's offsets;
/// behindDoc picks the z-order. Never moves the pen.
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

/// `layoutTextBox` — floating text boxes overlay the page at the current pen
/// without consuming height; inline ones flow like any other block.
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
        let incremental = layout_document_incremental(
            &mut incremental_input,
            &previous.layout,
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
