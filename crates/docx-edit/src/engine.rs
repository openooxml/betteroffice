//! Resident editor-engine state.

use std::borrow::Cow;
use std::cell::{Cell, RefCell};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::rc::Rc;

use docx_layout::display_list::DisplayList;
use docx_layout::footnotes::{
    FOOTNOTE_COLUMN_GAP_PX, OrderedMap, apply_note_presentation, assign_note_presentations,
    attach_note_areas, build_note_presentations, collect_note_refs, map_note_anchors_to_pages,
    map_notes_to_pages, stabilize_note_layout, stamp_note_pages,
};
use docx_layout::header_footer::{
    HeaderFooterKind, HeaderFooterMetrics, HeaderFooterPayload, HeaderFooterType,
    HeaderFooterVariant, extend_body_margins, measure_header_footer,
    resolve_header_footer_field_widths,
};
use docx_layout::hit::{CaretRect, VerticalDirection};
use docx_layout::paragraph_spacing::resolve_doc_grid_pitch;
use docx_layout::paragraph_spacing::resolve_line_unit_spacing;
use docx_layout::place::LayoutCheckpoint;
use docx_layout::regions::{
    DocumentRegions, RegionLayoutInput, apply_document_regions, apply_section_geometry,
    apply_section_geometry_to_blocks, effective_header_footer_refs,
};
use docx_layout::types::{
    BlockExtent, BlockId, ColumnLayout, Input as LayoutInput, Layout, LayoutBlock, MeasuredBlock,
    ParagraphExtent, Run,
};
use serde::Serialize;
use yrs::Subscription;

use crate::EditingDoc;
use crate::bridge::{BridgeError, RenderEnv, yrs_doc_to_layout_blocks};
use crate::frame_delta::{
    FrameEpochs, FramePageSnapshot, encode_frame_delta, encode_frame_delta_incremental,
};

#[derive(Debug)]
struct LoweredStory {
    doc_epoch: u64,
    env: RenderEnv,
    /// Shared so a reader can hold the lowering it asked for without the cache
    /// borrow, and without copying the story.
    blocks: Rc<Vec<LayoutBlock>>,
    /// Lazily serialized layout blocks.
    serialized_blocks: Option<String>,
}

#[derive(Debug, Default)]
struct RenderState {
    stories: HashMap<String, LoweredStory>,
    cache_hits: u64,
    cache_misses: u64,
}

#[derive(Debug)]
struct MeasureTemplate {
    envelope: serde_json::Value,
    resident_safe: bool,
}

#[derive(Debug, Default)]
struct MeasurementState {
    templates: HashMap<String, MeasureTemplate>,
    compatibility_calls: u64,
    resident_measure_calls: u64,
    resident_reused_blocks: u64,
}

#[derive(Debug)]
struct ResidentRegionState {
    request_json: String,
    headers_footers: Option<serde_json::Value>,
    /// Inputs retained from the last full region pass so a plain body-text
    /// edit can relayout residently. `None` when the pass was not
    /// resident-body or the document shape rules the fast path out.
    fast_path: Option<RegionFastPathState>,
}

/// Retained region-pass configuration consumed by
/// [`EngineSession::apply_and_layout_regions_resident`]. `Rc` keeps the
/// per-keystroke handoff clone-free.
#[derive(Debug)]
struct RegionFastPathState {
    regions: Rc<DocumentRegions>,
    measurement: Rc<docx_layout::measure_blocks::MeasurementConfig>,
    /// `measurement` hashed once so a later full pass can verify the retained
    /// arena was measured under the same config without re-serializing.
    measurement_fingerprint: u64,
    /// True when the last full pass had no normal footnote/endnote contents
    /// and no note references — the fast path skips note stabilization
    /// entirely, so it requires a note-free document.
    notes_clear: bool,
}

#[derive(Debug)]
struct ResidentLayoutInput {
    input: LayoutInput,
    block_fingerprints: Vec<u64>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct RegionLayoutOutput<'a> {
    measured: &'a [MeasuredBlock],
    options: &'a docx_layout::types::LayoutOptions,
    layout: &'a Layout,
    #[serde(skip_serializing_if = "Option::is_none")]
    headers_footers: Option<&'a serde_json::Value>,
    notes_converged: bool,
}

fn serialize_region_layout(
    input: &LayoutInput,
    layout: &Layout,
    headers_footers: Option<&serde_json::Value>,
    notes_converged: bool,
) -> Result<String, String> {
    serde_json::to_string(&RegionLayoutOutput {
        measured: &input.measured,
        options: &input.options,
        layout,
        headers_footers,
        notes_converged,
    })
    .map_err(|error| format!("serialize: {error}"))
}

fn reservation_options(reserved: &OrderedMap<u32, f64>) -> Option<BTreeMap<String, f64>> {
    (!reserved.is_empty()).then(|| {
        reserved
            .iter()
            .map(|(page, height)| (page.to_string(), *height))
            .collect()
    })
}

fn layout_error_message(error: docx_layout::LayoutError) -> String {
    match error {
        docx_layout::LayoutError::Unsupported(_) => "UNSUPPORTED".to_owned(),
        docx_layout::LayoutError::Invalid(reason) => reason,
    }
}

fn column_measurement_width(
    page_width: f64,
    margins: &docx_layout::types::PageMargins,
    columns: Option<&ColumnLayout>,
) -> f64 {
    let content_width = (page_width - margins.left - margins.right).max(1.0);
    let Some(columns) = columns.filter(|columns| columns.count > 1.0) else {
        return content_width;
    };
    if columns.equal_width == Some(false)
        && let Some(width) = columns
            .columns
            .as_ref()
            .and_then(|authored| authored.first())
            .and_then(|column| column.width)
            .filter(|width| width.is_finite() && *width > 0.0)
    {
        return width;
    }
    ((content_width - (columns.count - 1.0) * columns.gap) / columns.count).floor()
}

fn region_measurement_widths<'a>(
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
    input: &LayoutInput,
    regions: &DocumentRegions,
) -> Vec<f64> {
    let fallback_size = input
        .options
        .page_size
        .clone()
        .unwrap_or(docx_layout::types::Size {
            w: 816.0,
            h: 1056.0,
        });
    let fallback_margins =
        docx_layout::section_breaks::resolve_page_margins(input.options.margins.as_ref());
    let mut section_index = 0;
    blocks
        .into_iter()
        .map(|block| {
            let section = regions
                .sections
                .get(section_index)
                .or_else(|| regions.sections.last());
            let size = section
                .and_then(|section| section.page_size.as_ref())
                .unwrap_or(&fallback_size);
            let margins = section
                .and_then(|section| section.margins.as_ref())
                .unwrap_or(&fallback_margins);
            let columns = section
                .and_then(|section| section.columns.as_ref())
                .or(input.options.columns.as_ref());
            let width = column_measurement_width(size.w, margins, columns);
            if matches!(block, LayoutBlock::SectionBreak(_)) {
                section_index += 1;
            }
            width
        })
        .collect()
}

fn initial_float_page_geometry(
    input: &LayoutInput,
    regions: &DocumentRegions,
) -> docx_layout::measure_blocks::FloatPageGeometry {
    let section = regions.sections.first();
    let size = section
        .and_then(|section| section.page_size.clone())
        .or_else(|| input.options.page_size.clone())
        .unwrap_or(docx_layout::types::Size {
            w: 816.0,
            h: 1056.0,
        });
    let margins = docx_layout::section_breaks::resolve_page_margins(
        section
            .and_then(|section| section.margins.as_ref())
            .or(input.options.margins.as_ref()),
    );
    docx_layout::measure_blocks::FloatPageGeometry {
        page_width: size.w,
        margin_left: margins.left,
        page_height: size.h,
        margin_top: margins.top,
        content_height: size.h - margins.top - margins.bottom,
    }
}

fn stabilize_shape_wrapping(
    input: &mut LayoutInput,
    regions: &DocumentRegions,
    measurement: &docx_layout::measure_blocks::MeasurementConfig,
) -> Result<bool, docx_layout::LayoutError> {
    let shapes = input
        .measured
        .iter()
        .enumerate()
        .filter_map(|(index, measured)| {
            let LayoutBlock::Shape(shape) = &measured.block else {
                return None;
            };
            let horizontal = shape.position.as_ref()?.horizontal.as_ref()?;
            (matches!(
                shape.wrap_type.as_deref(),
                Some("square" | "tight" | "through" | "topAndBottom")
            ) && (horizontal.align.as_deref() == Some("inside")
                || matches!(
                    horizontal.relative_to.as_deref(),
                    Some("insideMargin" | "outsideMargin")
                )))
            .then(|| (index, shape.id.clone()))
        })
        .collect::<Vec<_>>();
    if shapes.is_empty() {
        return Ok(false);
    }
    let mut blocks = input
        .measured
        .iter()
        .map(|measured| measured.block.clone())
        .collect::<Vec<_>>();
    let widths = region_measurement_widths(blocks.iter(), input, regions);
    let geometry = initial_float_page_geometry(input, regions);
    let mut previous_offsets = BTreeMap::new();
    let mut touched = false;
    for _ in 0..shapes.len() + 2 {
        let layout = docx_layout::place::layout_document(input)?;
        let offsets = shapes
            .iter()
            .filter_map(|(index, id)| {
                layout
                    .pages
                    .iter()
                    .flat_map(|page| &page.fragments)
                    .find_map(|fragment| {
                        let docx_layout::types::Fragment::Shape(shape) = fragment else {
                            return None;
                        };
                        (shape.block_id == *id)
                            .then_some(shape.wrap_offset_x)
                            .flatten()
                            .map(|x| (*index, x))
                    })
            })
            .collect::<BTreeMap<_, _>>();
        if offsets == previous_offsets {
            return Ok(touched);
        }
        let measures = docx_layout::measure_blocks::measure_blocks_with_shape_offsets(
            &mut blocks,
            &widths,
            measurement,
            Some(&geometry),
            &offsets,
        )
        .map_err(docx_layout::LayoutError::Invalid)?;
        for (measured, (block, measure)) in
            input.measured.iter_mut().zip(blocks.iter().zip(measures))
        {
            measured.block = block.clone();
            measured.measure = measure;
        }
        touched = true;
        previous_offsets = offsets;
    }
    Err(docx_layout::LayoutError::Invalid(
        "anchored shape wrapping did not converge".to_owned(),
    ))
}

fn extend_input_for_header_footer(
    input: &mut LayoutInput,
    regions: &DocumentRegions,
    variants: &[HeaderFooterVariant],
) {
    if regions.sections.is_empty() {
        return;
    }
    let fallback_size = input
        .options
        .page_size
        .clone()
        .unwrap_or(docx_layout::types::Size {
            w: 816.0,
            h: 1056.0,
        });
    let fallback_margins =
        docx_layout::section_breaks::resolve_page_margins(input.options.margins.as_ref());
    let extended: Vec<_> = regions
        .sections
        .iter()
        .enumerate()
        .map(|(section_index, section)| {
            let page_size = section
                .page_size
                .clone()
                .unwrap_or_else(|| fallback_size.clone());
            let requested = section.margins.as_ref().or(input.options.margins.as_ref());
            let margins = match requested {
                Some(signed) => {
                    let mut base = signed.clone();
                    if base.header.is_none() {
                        base.header = Some(signed.top.abs());
                    }
                    if base.footer.is_none() {
                        base.footer = Some(signed.bottom.abs());
                    }
                    base
                }
                None => fallback_margins.clone(),
            };
            let header_height = variants
                .iter()
                .filter(|variant| {
                    variant.section_index == section_index
                        && variant.kind == HeaderFooterKind::Header
                })
                .map(|variant| variant.flow_height)
                .fold(0.0_f64, f64::max);
            let footer_height = variants
                .iter()
                .filter(|variant| {
                    variant.section_index == section_index
                        && variant.kind == HeaderFooterKind::Footer
                })
                .map(|variant| variant.flow_height)
                .fold(0.0_f64, f64::max);
            extend_body_margins(&page_size, &margins, header_height, footer_height)
        })
        .collect();
    input.options.margins = extended.first().cloned();
    input.options.final_margins = extended.last().cloned();
    let mut section_index = 0;
    for measured in &mut input.measured {
        let LayoutBlock::SectionBreak(section_break) = &mut measured.block else {
            continue;
        };
        if let Some(margins) = extended.get(section_index) {
            section_break.margins = Some(margins.clone());
        }
        section_index += 1;
    }
}

#[derive(Debug, Default)]
struct PaginationState {
    input: Option<LayoutInput>,
    /// Fingerprint of the measurement config that produced the retained
    /// `input` arena; the region path only reuses extents measured under an
    /// identical config.
    measured_with: Option<u64>,
    layout: Option<Layout>,
    checkpoints: Vec<LayoutCheckpoint>,
    block_fingerprints: Vec<u64>,
    options_fingerprint: u64,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: HashMap<String, i64>,
    last_incremental: bool,
    layout_epoch: u64,
    pagination_calls: u64,
    incremental_pagination_calls: u64,
    pagination_blocks_placed: u64,
}

#[derive(Debug, Default)]
struct DisplayState {
    list: Option<DisplayList>,
    resident_input: Option<docx_layout::display_list::ResidentDisplayInput>,
    frame_epoch: u64,
    display_builds: u64,
    binary_frame_epoch: u64,
    pages: Vec<FramePageSnapshot>,
    next_page_id: u64,
    extras_fingerprint: u64,
    extras_json: Option<String>,
    incremental_display_builds: u64,
    rebuilt_display_pages: u64,
}

/// Engine observability snapshot.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EngineStats {
    pub doc_epoch: u64,
    pub lowered_story_count: usize,
    pub lowered_block_count: usize,
    pub lower_cache_hits: u64,
    pub lower_cache_misses: u64,
    pub retained_measure_templates: usize,
    pub compatibility_measure_calls: u64,
    pub resident_measure_calls: u64,
    pub resident_reused_blocks: u64,
    pub layout_epoch: u64,
    pub retained_measured_blocks: usize,
    pub retained_pages: usize,
    pub pagination_calls: u64,
    pub incremental_pagination_calls: u64,
    pub pagination_blocks_placed: u64,
    pub retained_checkpoints: usize,
    pub rebuilt_pages: usize,
    pub frame_epoch: u64,
    pub retained_display_pages: usize,
    pub retained_display_primitives: usize,
    pub display_builds: u64,
    pub incremental_display_builds: u64,
    pub rebuilt_display_pages: u64,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ResidentCaretRect {
    pub page_index: usize,
    pub page_id: String,
    pub x: f64,
    pub y: f64,
    pub height: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ResidentCaretSnapshot {
    pub frame_epoch: u64,
    pub caret_rect: Option<ResidentCaretRect>,
}

/// Fine-grained timings for one profiled resident input transaction.
///
/// The engine accepts its clock from the wasm facade so native builds keep no
/// browser dependency and the ordinary (unprofiled) input path pays no timer
/// calls. Values are milliseconds.
#[derive(Clone, Copy, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EngineApplyProfile {
    pub lower_ms: f64,
    pub measure_ms: f64,
    pub paginate_ms: f64,
    pub display_input_ms: f64,
    pub display_build_ms: f64,
    pub display_finalize_ms: f64,
    pub display_ms: f64,
    pub encode_ms: f64,
}

/// Stage boundaries emitted by the resident region fast path so the profiler
/// can attribute lower/measure time separately from pagination.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RegionResidentPhase {
    Lowered,
    Measured,
}

/// Long-lived owner of the authoritative editing document and its retained
/// render projections.
///
/// The yrs update observer advances `doc_epoch` for every committed local or
/// remote transaction. Render caches are generation-tagged instead of being
/// eagerly cleared, so an in-flight read can never publish blocks from a
/// different document generation.
pub struct EngineSession {
    doc: EditingDoc,
    doc_epoch: Rc<Cell<u64>>,
    // Kept alive for the lifetime of the document. Dropping it unregisters the
    // observer before the Rc epoch source is released.
    _doc_epoch_observer: Subscription,
    render: RefCell<RenderState>,
    measurement: RefCell<MeasurementState>,
    regions: RefCell<Option<ResidentRegionState>>,
    pagination: RefCell<PaginationState>,
    display: RefCell<DisplayState>,
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

fn strip_absolute_positions(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Array(values) => {
            for value in values {
                strip_absolute_positions(value);
            }
        }
        serde_json::Value::Object(fields) => {
            for key in ["pmStart", "pmEnd", "docStart", "docEnd"] {
                fields.remove(key);
            }
            for value in fields.values_mut() {
                strip_absolute_positions(value);
            }
        }
        _ => {}
    }
}

fn measured_fingerprint(measured: &MeasuredBlock) -> Result<u64, String> {
    let mut value: serde_json::Value = serde_json::to_string(measured)
        .map_err(|error| format!("fingerprint measured block: {error}"))
        .and_then(|s| {
            serde_json::from_str(&s).map_err(|error| format!("fingerprint measured block: {error}"))
        })?;
    strip_absolute_positions(&mut value);
    serde_json::to_vec(&value)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint measured block: {error}"))
}

fn measured_fingerprints(input: &LayoutInput) -> Result<Vec<u64>, String> {
    input.measured.iter().map(measured_fingerprint).collect()
}

fn value_block_key(value: &serde_json::Value) -> Option<String> {
    let id = value.get("block")?.get("id")?;
    match id {
        serde_json::Value::String(value) => Some(value.clone()),
        serde_json::Value::Number(value) => Some(value.to_string()),
        _ => None,
    }
}

fn measure_template_is_resident_safe(value: &serde_json::Value) -> bool {
    value
        .get("floatingZones")
        .and_then(serde_json::Value::as_array)
        .is_none_or(Vec::is_empty)
        && value
            .get("paragraphYOffset")
            .and_then(serde_json::Value::as_f64)
            .is_none_or(|offset| offset == 0.0)
}

fn options_fingerprint(input: &LayoutInput) -> Result<u64, String> {
    serde_json::to_vec(&input.options)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint layout options: {error}"))
}

fn block_key(id: &BlockId) -> Cow<'_, str> {
    match id {
        BlockId::Num(value) if value.fract() == 0.0 => Cow::Owned(format!("{}", *value as i64)),
        BlockId::Num(value) => Cow::Owned(value.to_string()),
        BlockId::Str(value) => Cow::Borrowed(value),
    }
}

fn paragraph_identity(block: &LayoutBlock) -> Option<(&BlockId, Option<f64>)> {
    match block {
        LayoutBlock::Paragraph(paragraph)
            if paragraph
                .runs
                .iter()
                .all(|run| !matches!(run, Run::Image(_) | Run::Unsupported)) =>
        {
            Some((&paragraph.id, paragraph.pm_start))
        }
        _ => None,
    }
}

fn resident_block_slots_match(previous: &LayoutBlock, next: &LayoutBlock) -> bool {
    match (paragraph_identity(previous), paragraph_identity(next)) {
        (Some((previous_id, _)), Some((next_id, _))) => {
            block_key(previous_id) == block_key(next_id)
        }
        (None, None) => true,
        _ => false,
    }
}

fn fragment_identity(block: &LayoutBlock) -> Option<&BlockId> {
    match block {
        LayoutBlock::Paragraph(block) => Some(&block.id),
        LayoutBlock::Table(block) => Some(&block.id),
        LayoutBlock::Image(block) => Some(&block.id),
        LayoutBlock::TextBox(block) => Some(&block.id),
        LayoutBlock::Shape(block) => Some(&block.id),
        LayoutBlock::Chart(block) => Some(&block.id),
        _ => None,
    }
}

fn resident_fragment_keys_match(previous: &LayoutBlock, next: &LayoutBlock) -> bool {
    match (fragment_identity(previous), fragment_identity(next)) {
        (Some(previous_id), Some(next_id)) => block_key(previous_id) == block_key(next_id),
        (None, None) => true,
        _ => false,
    }
}

fn effective_paragraph_style(paragraph: &docx_layout::types::ParagraphBlock) -> &str {
    let Some(attrs) = paragraph.attrs.as_ref() else {
        return "";
    };
    attrs
        .effective_style_id
        .as_deref()
        .filter(|style| !style.is_empty())
        .or_else(|| attrs.style_id.as_deref().filter(|style| !style.is_empty()))
        .unwrap_or("")
}

fn empty_paragraph_runs(paragraph: &docx_layout::types::ParagraphBlock) -> bool {
    paragraph.runs.is_empty()
        || matches!(paragraph.runs.as_slice(), [Run::Text(run)] if run.text.is_empty())
}

fn contextual_spacing_enabled(paragraph: &docx_layout::types::ParagraphBlock) -> bool {
    paragraph
        .attrs
        .as_ref()
        .is_some_and(|attrs| attrs.contextual_spacing.unwrap_or(false))
}

/// First paragraph of a table's leading cell, used by the empty-paragraph +
/// table arm of the contextual-spacing rule.
fn first_cell_paragraph(
    table: &docx_layout::types::TableBlock,
) -> Option<&docx_layout::types::ParagraphBlock> {
    table
        .rows
        .first()
        .and_then(|row| row.cells.first())
        .and_then(|cell| cell.blocks.first())
        .and_then(|block| match block {
            LayoutBlock::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
}

/// Put extents moved out of the retained arena back; consumed in index order.
fn restore_moved_measures(measured: &mut [MeasuredBlock], entries: Vec<MeasuredBlock>) {
    for (consumed, entry) in entries.into_iter().enumerate() {
        measured[consumed].measure = entry.measure;
    }
}

/// Mirror `contextual_spacing_pair`'s writes for a freshly lowered `owned`
/// block before it can compare equal to a retained one: `before` vs the
/// previous sibling and `after` vs the next (the table arm uses the first
/// cell paragraph for the style check only — canonical writes `after` on the
/// empty paragraph and never touches the table's `before`).
fn suppress_contextual_spacing(blocks: &[LayoutBlock], index: usize, owned: &mut LayoutBlock) {
    if let Some(LayoutBlock::Paragraph(previous)) = index.checked_sub(1).and_then(|i| blocks.get(i))
    {
        let suppress_before = match &*owned {
            LayoutBlock::Paragraph(paragraph) => {
                effective_paragraph_style(paragraph) == effective_paragraph_style(previous)
                    && contextual_spacing_enabled(paragraph)
            }
            _ => false,
        };
        if suppress_before
            && let LayoutBlock::Paragraph(paragraph) = owned
            && let Some(spacing) = paragraph
                .attrs
                .as_mut()
                .and_then(|attrs| attrs.spacing.as_mut())
        {
            spacing.before = Some(0.0);
        }
    }
    let LayoutBlock::Paragraph(current) = owned else {
        return;
    };
    let suppress_after = match blocks.get(index + 1) {
        Some(LayoutBlock::Paragraph(next)) => {
            effective_paragraph_style(next) == effective_paragraph_style(current)
                && contextual_spacing_enabled(current)
        }
        Some(LayoutBlock::Table(table))
            if table.floating.is_none() && empty_paragraph_runs(current) =>
        {
            first_cell_paragraph(table).is_some_and(|inner| {
                effective_paragraph_style(inner) == effective_paragraph_style(current)
                    && contextual_spacing_enabled(current)
            })
        }
        _ => false,
    };
    if suppress_after
        && let Some(spacing) = current
            .attrs
            .as_mut()
            .and_then(|attrs| attrs.spacing.as_mut())
    {
        spacing.after = Some(0.0);
    }
}

fn nested_block_keys(block: &LayoutBlock, out: &mut Vec<String>) {
    match block {
        LayoutBlock::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for nested in &cell.blocks {
                        if let Some(id) = nested.block_id() {
                            out.push(block_key(id).into_owned());
                        }
                        nested_block_keys(nested, out);
                    }
                }
            }
        }
        LayoutBlock::TextBox(textbox) => {
            for paragraph in &textbox.content {
                out.push(block_key(&paragraph.id).into_owned());
            }
        }
        LayoutBlock::Shape(shape) => {
            if let Some(inner_text) = &shape.inner_text {
                for paragraph in inner_text {
                    out.push(block_key(&paragraph.id).into_owned());
                }
            }
            for child in &shape.children {
                if let Some(inner_text) = &child.inner_text {
                    for paragraph in inner_text {
                        out.push(block_key(&paragraph.id).into_owned());
                    }
                }
            }
        }
        _ => {}
    }
}

fn position_deltas(previous: &LayoutInput, next: &LayoutInput) -> HashMap<String, i64> {
    previous
        .measured
        .iter()
        .zip(&next.measured)
        .filter_map(|(previous, next)| {
            let previous_id = previous.block.block_id()?;
            let next_id = next.block.block_id()?;
            let key = block_key(previous_id);
            if key != block_key(next_id) {
                return None;
            }
            let previous_start = previous.block.pm_start();
            let next_start = next.block.pm_start();
            let delta = next_start? as i64 - previous_start? as i64;
            if delta == 0 {
                return None;
            }
            let mut entries = vec![(key.into_owned(), delta)];
            // Nested paragraphs (table cells, text boxes, shape text) carry
            // their own ids on display primitives but shift with the
            // enclosing block.
            let mut nested = Vec::new();
            nested_block_keys(&next.block, &mut nested);
            entries.extend(nested.into_iter().map(|key| (key, delta)));
            Some(entries)
        })
        .flatten()
        .collect()
}

fn incremental_eligible(
    previous: &PaginationState,
    next: &LayoutInput,
    next_options_fingerprint: u64,
) -> bool {
    let Some(previous_input) = previous.input.as_ref() else {
        return false;
    };
    previous.layout.is_some()
        && !previous.checkpoints.is_empty()
        && previous.options_fingerprint == next_options_fingerprint
        && previous_input.measured.len() == next.measured.len()
        && next
            .options
            .footnote_reserved_heights
            .as_ref()
            .is_none_or(|heights| heights.is_empty())
        && next
            .options
            .columns
            .as_ref()
            .is_none_or(|columns| columns.count <= 1.0)
        && previous_input
            .measured
            .iter()
            .zip(&next.measured)
            .all(|(previous, next)| resident_fragment_keys_match(&previous.block, &next.block))
}

impl EngineSession {
    pub fn new(client_id: u64) -> Self {
        let doc = EditingDoc::new(client_id);
        let doc_epoch = Rc::new(Cell::new(0_u64));
        let observer_epoch = Rc::clone(&doc_epoch);
        let observer = doc
            .yrs_doc()
            .observe_update_v1(move |_txn, _event| {
                observer_epoch.set(observer_epoch.get().wrapping_add(1));
            })
            .expect("EngineSession document update observer registers");
        Self {
            doc,
            doc_epoch,
            _doc_epoch_observer: observer,
            render: RefCell::new(RenderState::default()),
            measurement: RefCell::new(MeasurementState::default()),
            regions: RefCell::new(None),
            pagination: RefCell::new(PaginationState::default()),
            display: RefCell::new(DisplayState::default()),
        }
    }

    /// Editing document.
    pub fn doc(&self) -> &EditingDoc {
        &self.doc
    }

    pub fn doc_epoch(&self) -> u64 {
        self.doc_epoch.get()
    }

    /// Runs a callback with resident lowered blocks.
    pub fn with_lowered_story<T>(
        &self,
        story: &str,
        env: &RenderEnv,
        read: impl FnOnce(&[LayoutBlock]) -> T,
    ) -> Result<T, BridgeError> {
        self.with_lowered_story_observed(story, env, &mut || {}, read)
    }

    /// Whether `story` is already lowered for this document epoch and
    /// environment.
    fn story_is_resident(&self, story: &str, epoch: u64, env: &RenderEnv) -> bool {
        self.render
            .borrow()
            .stories
            .get(story)
            .is_some_and(|cached| cached.doc_epoch == epoch && cached.env == *env)
    }

    fn lower_story_into_cache(
        &self,
        story: &str,
        epoch: u64,
        env: &RenderEnv,
    ) -> Result<(), BridgeError> {
        let blocks = yrs_doc_to_layout_blocks(&self.doc, story, env)?;
        let mut render = self.render.borrow_mut();
        render.cache_misses = render.cache_misses.wrapping_add(1);
        render.stories.insert(
            story.to_owned(),
            LoweredStory {
                doc_epoch: epoch,
                env: env.clone(),
                blocks: Rc::new(blocks),
                serialized_blocks: None,
            },
        );
        Ok(())
    }

    /// [`Self::with_lowered_story`] with a hook between lowering and the read.
    /// The lowering is claimed before `after_lower` fires, so an observer
    /// supplied by the host may re-enter the engine — even re-lowering this
    /// story — without disturbing what the caller reads.
    fn with_lowered_story_observed<T>(
        &self,
        story: &str,
        env: &RenderEnv,
        after_lower: &mut dyn FnMut(),
        read: impl FnOnce(&[LayoutBlock]) -> T,
    ) -> Result<T, BridgeError> {
        let epoch = self.doc_epoch();
        if self.story_is_resident(story, epoch, env) {
            let mut render = self.render.borrow_mut();
            render.cache_hits = render.cache_hits.wrapping_add(1);
        } else {
            self.lower_story_into_cache(story, epoch, env)?;
        }
        let blocks = Rc::clone(
            &self
                .render
                .borrow()
                .stories
                .get(story)
                .expect("resident story exists after lowering")
                .blocks,
        );
        after_lower();
        Ok(read(&blocks))
    }

    /// Serializes resident lowered blocks.
    pub fn lower_story_json(&self, story: &str, env: &RenderEnv) -> Result<String, BridgeError> {
        self.with_lowered_story(story, env, |_| ())?;
        let mut render = self.render.borrow_mut();
        let lowered = render
            .stories
            .get_mut(story)
            .expect("resident story exists after lowering");
        Ok(lowered
            .serialized_blocks
            .get_or_insert_with(|| {
                serde_json::to_string(&lowered.blocks)
                    .expect("LayoutBlock serialization is infallible after lowering")
            })
            .clone())
    }

    pub fn stats(&self) -> EngineStats {
        let render = self.render.borrow();
        let measurement = self.measurement.borrow();
        let pagination = self.pagination.borrow();
        let display = self.display.borrow();
        EngineStats {
            doc_epoch: self.doc_epoch(),
            lowered_story_count: render.stories.len(),
            lowered_block_count: render
                .stories
                .values()
                .map(|story| story.blocks.len())
                .sum(),
            lower_cache_hits: render.cache_hits,
            lower_cache_misses: render.cache_misses,
            retained_measure_templates: measurement.templates.len(),
            compatibility_measure_calls: measurement.compatibility_calls,
            resident_measure_calls: measurement.resident_measure_calls,
            resident_reused_blocks: measurement.resident_reused_blocks,
            layout_epoch: pagination.layout_epoch,
            retained_measured_blocks: pagination
                .input
                .as_ref()
                .map_or(0, |input| input.measured.len()),
            retained_pages: pagination
                .layout
                .as_ref()
                .map_or(0, |layout| layout.pages.len()),
            pagination_calls: pagination.pagination_calls,
            incremental_pagination_calls: pagination.incremental_pagination_calls,
            pagination_blocks_placed: pagination.pagination_blocks_placed,
            retained_checkpoints: pagination.checkpoints.len(),
            rebuilt_pages: pagination
                .rebuilt_page_end
                .saturating_sub(pagination.rebuilt_page_start),
            frame_epoch: display.frame_epoch,
            retained_display_pages: display.list.as_ref().map_or(0, |list| list.pages.len()),
            retained_display_primitives: display.list.as_ref().map_or(0, |list| {
                list.pages.iter().map(|page| page.primitives.len()).sum()
            }),
            display_builds: display.display_builds,
            incremental_display_builds: display.incremental_display_builds,
            rebuilt_display_pages: display.rebuilt_display_pages,
        }
    }

    /// Compatibility paragraph measurement which also records the immutable
    /// width/font/compatibility envelope under the paragraph's stable block
    /// id. A later resident edit replaces only `block` and reuses the rest of
    /// this envelope, so the host no longer orchestrates dirty measurement.
    pub fn measure_paragraph_json(&self, input_json: &str) -> Result<String, String> {
        let value: serde_json::Value =
            serde_json::from_str(input_json).map_err(|error| format!("invalid: parse: {error}"))?;
        let key = value_block_key(&value)
            .ok_or_else(|| "invalid: measurement block requires a stable id".to_owned())?;
        let output = docx_layout::measure_paragraph_json_resident(input_json)?;
        let resident_safe = measure_template_is_resident_safe(&value);
        let mut measurement = self.measurement.borrow_mut();
        measurement.templates.insert(
            key,
            MeasureTemplate {
                envelope: value,
                resident_safe,
            },
        );
        measurement.compatibility_calls = measurement.compatibility_calls.wrapping_add(1);
        Ok(output)
    }

    /// Invalidate paragraph templates when font ids are reset. Existing
    /// layout/display state remains a valid painted snapshot, but a new edit
    /// must pass through the compatibility readiness/layout path first.
    pub fn clear_measurement_templates(&self) {
        self.measurement.borrow_mut().templates.clear();
    }

    fn measurement_envelope_for_block(
        &self,
        key: &str,
        previous_block: &LayoutBlock,
    ) -> Option<serde_json::Value> {
        let measurement = self.measurement.borrow();
        if let Some(template) = measurement.templates.get(key)
            && template.resident_safe
        {
            return Some(template.envelope.clone());
        }
        measurement.templates.values().find_map(|template| {
            if !template.resident_safe {
                return None;
            }
            let block: LayoutBlock = serde_json::to_string(template.envelope.get("block")?)
                .ok()
                .and_then(|s| serde_json::from_str(&s).ok())?;
            (block == *previous_block).then(|| template.envelope.clone())
        })
    }

    /// Parses, paginates, and retains measured input and layout.
    pub fn layout_document_json(&self, input_json: &str) -> Result<String, String> {
        let input: LayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        self.regions.borrow_mut().take();
        self.layout_document_value(input)?;
        let pagination = self.pagination.borrow();
        serde_json::to_string(
            pagination
                .layout
                .as_ref()
                .expect("layout retained after successful pagination"),
        )
        .map_err(|error| format!("serialize: {error}"))
    }

    pub fn layout_font_requirements_json(&self, input_json: &str) -> Result<String, String> {
        let request: RegionLayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        let (input, regions, notes, _, render_env, body_story) = request.split();
        let mut blocks = input
            .measured
            .into_iter()
            .map(|measured| measured.block)
            .collect::<Vec<_>>();
        if let Some(body_story) = body_story {
            let render_env: RenderEnv = serde_json::to_string(&render_env)
                .ok()
                .and_then(|s| serde_json::from_str(&s).ok())
                .ok_or_else(|| "parse render environment".to_owned())?;
            let mut stories = BTreeSet::from([body_story]);
            for section_index in 0..regions.sections.len() {
                let Some(refs) = effective_header_footer_refs(&regions, section_index) else {
                    continue;
                };
                stories.extend(
                    [
                        refs.header_default,
                        refs.header_first,
                        refs.header_even,
                        refs.footer_default,
                        refs.footer_first,
                        refs.footer_even,
                    ]
                    .into_iter()
                    .flatten()
                    .map(|r_id| format!("hf:{r_id}")),
                );
            }
            stories.extend(notes.contents.into_iter().map(|content| {
                let prefix = match content.note_kind {
                    docx_layout::footnotes::NoteKind::Footnote => "fn",
                    docx_layout::footnotes::NoteKind::Endnote => "en",
                };
                format!("{prefix}:{}", content.id)
            }));
            for story in stories {
                let mut story_blocks = self
                    .with_lowered_story(&story, &render_env, <[LayoutBlock]>::to_vec)
                    .map_err(|error| error.to_string())?;
                blocks.append(&mut story_blocks);
            }
        }
        serde_json::to_string(&docx_layout::measure_blocks::collect_font_requirements(
            &blocks,
        ))
        .map_err(|error| format!("serialize: {error}"))
    }

    /// Full-document pagination with section/page region orchestration owned
    /// by the resident engine. The returned envelope is ready for the Rust
    /// display-list builder without host-side layout mutation.
    pub fn layout_document_with_regions_json(&self, input_json: &str) -> Result<String, String> {
        let notes_converged = self.layout_document_with_regions_value(input_json)?;
        let pagination = self.pagination.borrow();
        let regions_state = self.regions.borrow();
        let state = regions_state
            .as_ref()
            .expect("region state retained after successful region layout");
        serialize_region_layout(
            pagination
                .input
                .as_ref()
                .expect("input retained after successful pagination"),
            pagination
                .layout
                .as_ref()
                .expect("layout retained after successful pagination"),
            state.headers_footers.as_ref(),
            notes_converged,
        )
    }

    /// Full region pass whose reply omits the measured arena — tens of MB of
    /// shaping data a worker-rendered host never reads. Fetch the retained
    /// inputs on demand via [`Self::retained_kernel_inputs_json`].
    pub fn layout_document_with_regions_retained_json(
        &self,
        input_json: &str,
    ) -> Result<String, String> {
        let notes_converged = self.layout_document_with_regions_value(input_json)?;
        let pagination = self.pagination.borrow();
        let regions_state = self.regions.borrow();
        let state = regions_state
            .as_ref()
            .expect("region state retained after successful region layout");
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct RetainedRegionLayoutOutput<'a> {
            layout: &'a Layout,
            #[serde(skip_serializing_if = "Option::is_none")]
            headers_footers: Option<&'a serde_json::Value>,
            notes_converged: bool,
        }
        serde_json::to_string(&RetainedRegionLayoutOutput {
            layout: pagination
                .layout
                .as_ref()
                .expect("layout retained after successful pagination"),
            headers_footers: state.headers_footers.as_ref(),
            notes_converged,
        })
        .map_err(|error| format!("serialize: {error}"))
    }

    /// The retained measured arena and layout options, for a host taking the
    /// main-thread display-list fallback after a retained-only region layout.
    pub fn retained_kernel_inputs_json(&self) -> Result<String, String> {
        let pagination = self.pagination.borrow();
        let input = pagination
            .input
            .as_ref()
            .ok_or_else(|| "resident pagination input is not built".to_owned())?;
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct RetainedKernelInputs<'a> {
            measured: &'a [MeasuredBlock],
            options: &'a docx_layout::types::LayoutOptions,
        }
        serde_json::to_string(&RetainedKernelInputs {
            measured: &input.measured,
            options: &input.options,
        })
        .map_err(|error| format!("serialize: {error}"))
    }

    /// The full region pass minus the JSON envelope: pagination, note
    /// stabilization, and header/footer measurement all land in retained
    /// state. `apply_input`'s fallback consumes this directly so a keystroke
    /// never serializes a layout nobody reads. Returns `notes_converged`.
    fn layout_document_with_regions_value(&self, input_json: &str) -> Result<bool, String> {
        let request: RegionLayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        let (mut input, regions, mut notes, measurement, render_env, body_story) = request.split();
        let mut parsed_render_env = if render_env.is_null() {
            None
        } else {
            Some(
                serde_json::to_string(&render_env)
                    .ok()
                    .and_then(|s| serde_json::from_str::<RenderEnv>(&s).ok())
                    .ok_or_else(|| "parse render environment".to_owned())?,
            )
        };
        if regions.sections.len() <= 1
            && let Some(env) = &mut parsed_render_env
        {
            let line_px = regions.paragraph_spacing_line_px(0);
            env.paragraph_spacing_line_px = (line_px != 16.0).then_some(line_px);
            env.doc_grid_pitch_px = regions.doc_grid_snap_pitch_px(0);
        }
        let measurement_fingerprint = serde_json::to_vec(&measurement)
            .map(|bytes| hash_bytes(&bytes))
            .map_err(|error| format!("fingerprint measurement config: {error}"))?;
        let resident_body = body_story.is_some();
        let mut block_fingerprints: Option<Vec<u64>> = None;
        if let Some(story) = body_story.as_deref() {
            let render_env = parsed_render_env
                .as_ref()
                .ok_or_else(|| "resident body layout requires a render environment".to_owned())?;
            enum Arena {
                Reused(Vec<MeasuredBlock>, Vec<u64>),
                Full(Vec<LayoutBlock>),
            }
            let arena = self
                .with_lowered_story(story, render_env, |blocks| -> Result<Arena, String> {
                    apply_section_geometry(&mut input, &regions);
                    let widths = region_measurement_widths(blocks.iter(), &input, &regions);
                    let geometry = initial_float_page_geometry(&input, &regions);
                    let default_width = widths.first().copied().unwrap_or(0.0);
                    // floating extents depend on flow position; re-measure all
                    if docx_layout::measure_blocks::has_floating_zones(
                        blocks,
                        default_width,
                        &measurement,
                        Some(&geometry),
                    )? {
                        return Ok(Arena::Full(blocks.to_vec()));
                    }
                    match self.resident_region_measured(
                        blocks,
                        &widths,
                        &regions,
                        &measurement,
                        measurement_fingerprint,
                    )? {
                        Some((measured, fingerprints)) => Ok(Arena::Reused(measured, fingerprints)),
                        None => Ok(Arena::Full(blocks.to_vec())),
                    }
                })
                .map_err(|error| error.to_string())??;
            match arena {
                Arena::Reused(measured, fingerprints) => {
                    input.measured = measured;
                    block_fingerprints = Some(fingerprints);
                    apply_section_geometry(&mut input, &regions);
                }
                Arena::Full(mut blocks) => {
                    let mut section_index = 0;
                    for block in &mut blocks {
                        resolve_line_unit_spacing(
                            block,
                            regions.paragraph_spacing_line_px(section_index),
                        );
                        resolve_doc_grid_pitch(
                            block,
                            regions.doc_grid_snap_pitch_px(section_index),
                        );
                        if matches!(block, LayoutBlock::SectionBreak(_)) {
                            section_index += 1;
                        }
                    }
                    apply_section_geometry_to_blocks(&mut blocks, &mut input.options, &regions);
                    let widths = region_measurement_widths(blocks.iter(), &input, &regions);
                    let geometry = initial_float_page_geometry(&input, &regions);
                    let measures = docx_layout::measure_blocks::measure_blocks_with_floats(
                        &mut blocks,
                        &widths,
                        &measurement,
                        Some(&geometry),
                    )?;
                    input.measured = blocks
                        .into_iter()
                        .zip(measures)
                        .map(|(block, measure)| MeasuredBlock { block, measure })
                        .collect();
                }
            }
        } else {
            apply_section_geometry(&mut input, &regions);
        }
        let mut measured_headers_footers = if let Some(render_env) = parsed_render_env.as_ref() {
            self.measure_header_footer_payload(&mut input, &regions, &measurement, render_env)?
        } else {
            None
        };
        if resident_body
            && stabilize_shape_wrapping(&mut input, &regions, &measurement)
                .map_err(layout_error_message)?
        {
            block_fingerprints = None;
        }
        let block_fingerprints = match block_fingerprints {
            Some(fingerprints) => fingerprints,
            None => measured_fingerprints(&input)?,
        };
        // The note fixpoint replays `base_input`; `input` itself moves into
        // pagination and is never cloned again.
        let base_input = input.clone();
        self.layout_document_value_with_fingerprints(input, block_fingerprints)?;
        let mut initial_layout = self
            .pagination
            .borrow_mut()
            .layout
            .take()
            .expect("layout retained after successful pagination");
        apply_document_regions(&mut initial_layout, &regions);
        let refs = base_input
            .measured
            .iter()
            .flat_map(|measured| collect_note_refs(std::slice::from_ref(&measured.block)))
            .collect::<Vec<_>>();
        let presentations = build_note_presentations(&refs, &initial_layout.pages, &regions);
        assign_note_presentations(&mut notes.contents, &presentations);
        if resident_body {
            self.measure_resident_notes(
                &mut notes.contents,
                &refs,
                &initial_layout,
                &regions,
                &measurement,
                parsed_render_env
                    .as_ref()
                    .expect("resident body required render environment"),
            )?;
        }
        let stabilized = stabilize_note_layout(
            |reserved| {
                let mut pass = base_input.clone();
                pass.options.footnote_reserved_heights = reservation_options(reserved);
                if resident_body {
                    stabilize_shape_wrapping(&mut pass, &regions, &measurement)?;
                }
                let mut layout = docx_layout::place::layout_document(&mut pass)?;
                apply_document_regions(&mut layout, &regions);
                Ok(layout)
            },
            &refs,
            &notes.contents,
            initial_layout,
            &regions,
        )
        .map_err(layout_error_message)?;
        let notes_converged = stabilized.converged;
        if !stabilized.reserved_heights.is_empty() {
            let mut final_input = base_input;
            final_input.options.footnote_reserved_heights =
                reservation_options(&stabilized.reserved_heights);
            if resident_body {
                stabilize_shape_wrapping(&mut final_input, &regions, &measurement)
                    .map_err(layout_error_message)?;
            }
            self.layout_document_value(final_input)?;
        } else {
            self.pagination.borrow_mut().layout = Some(stabilized.layout);
        }
        let mut pagination = self.pagination.borrow_mut();
        let layout = pagination
            .layout
            .as_mut()
            .expect("layout retained after successful pagination");
        apply_document_regions(layout, &regions);
        let page_note_map = map_notes_to_pages(&layout.pages, &refs, &regions);
        stamp_note_pages(layout, &page_note_map, &regions);
        attach_note_areas(layout, &page_note_map, &notes.contents, &regions);
        let measured_headers_footers = measured_headers_footers
            .as_mut()
            .map(|payload| {
                resolve_header_footer_field_widths(payload, layout, &measurement)?;
                serde_json::to_string(payload)
                    .ok()
                    .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
                    .ok_or_else(|| "serialize headers/footers".to_owned())
            })
            .transpose()?;
        let headers_footers = measured_headers_footers.or_else(|| regions.headers_footers.clone());
        let notes_clear = notes.contents.is_empty() && refs.is_empty();
        // Multi-section documents are excluded: an edit can move a section
        // boundary without changing the total page count, which changes
        // section-relative page labels and the PAGE/NUMPAGES field widths
        // baked into the retained headers/footers payload. With one section,
        // an unchanged page count implies unchanged labels.
        let single_section = regions.sections.len() <= 1;
        drop(pagination);
        self.regions.replace(Some(ResidentRegionState {
            request_json: input_json.to_owned(),
            headers_footers,
            fast_path: (resident_body && single_section).then(|| RegionFastPathState {
                regions: Rc::new(regions),
                measurement: Rc::new(measurement),
                measurement_fingerprint,
                notes_clear,
            }),
        }));
        // Only the region-measured arena may seed the next pass's reuse walk.
        self.pagination.borrow_mut().measured_with =
            resident_body.then_some(measurement_fingerprint);
        Ok(notes_converged)
    }

    fn measure_resident_notes(
        &self,
        contents: &mut [docx_layout::footnotes::NoteContent],
        refs: &[docx_layout::footnotes::NoteRefLocation],
        layout: &Layout,
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        render_env: &RenderEnv,
    ) -> Result<(), String> {
        let anchors = map_note_anchors_to_pages(&layout.pages, refs);
        for content in contents {
            let Some(page_number) = anchors
                .iter()
                .find_map(|(page, ids)| ids.contains(&content.map_id()).then_some(*page))
            else {
                continue;
            };
            let Some(page) = layout.pages.iter().find(|page| page.number == page_number) else {
                continue;
            };
            let columns = regions.footnote_columns(page.region_section_index);
            let content_width = page.size.w - page.margins.left - page.margins.right;
            let width = ((content_width
                - (columns.saturating_sub(1) as f64) * FOOTNOTE_COLUMN_GAP_PX)
                / columns as f64)
                .max(1.0);
            let prefix = match content.note_kind {
                docx_layout::footnotes::NoteKind::Footnote => "fn",
                docx_layout::footnotes::NoteKind::Endnote => "en",
            };
            let mut blocks = self
                .with_lowered_story(
                    &format!("{prefix}:{}", content.id),
                    render_env,
                    <[LayoutBlock]>::to_vec,
                )
                .map_err(|error| error.to_string())?;
            for block in &mut blocks {
                resolve_line_unit_spacing(
                    block,
                    regions.paragraph_spacing_line_px(page.region_section_index),
                );
                resolve_doc_grid_pitch(
                    block,
                    regions.doc_grid_snap_pitch_px(page.region_section_index),
                );
            }
            apply_note_presentation(
                &mut blocks,
                content.display_number.unwrap_or(1),
                content.display_label.as_deref().unwrap_or("1"),
            );
            let measures =
                docx_layout::measure_blocks::measure_blocks(&mut blocks, width, measurement)?;
            content.height = measures
                .iter()
                .map(docx_layout::measure_blocks::extent_height)
                .sum();
            content.blocks = blocks;
            content.measures = measures;
        }
        Ok(())
    }

    fn measure_header_footer_payload(
        &self,
        input: &mut LayoutInput,
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        render_env: &RenderEnv,
    ) -> Result<Option<HeaderFooterPayload>, String> {
        let mut variants = Vec::new();
        for section_index in 0..regions.sections.len() {
            let Some(refs) = effective_header_footer_refs(regions, section_index) else {
                continue;
            };
            let section = &regions.sections[section_index];
            let page_size = section
                .page_size
                .clone()
                .or_else(|| input.options.page_size.clone())
                .unwrap_or(docx_layout::types::Size {
                    w: 816.0,
                    h: 1056.0,
                });
            let margins = docx_layout::section_breaks::resolve_page_margins(
                section.margins.as_ref().or(input.options.margins.as_ref()),
            );
            let content_width = (page_size.w - margins.left - margins.right).max(1.0);
            let descriptors = [
                (
                    HeaderFooterKind::Header,
                    HeaderFooterType::Default,
                    refs.header_default.clone().or_else(|| {
                        (!section.title_pg)
                            .then_some(refs.header_first.clone())
                            .flatten()
                    }),
                ),
                (
                    HeaderFooterKind::Header,
                    HeaderFooterType::First,
                    section.title_pg.then_some(refs.header_first).flatten(),
                ),
                (
                    HeaderFooterKind::Header,
                    HeaderFooterType::Even,
                    refs.header_even,
                ),
                (
                    HeaderFooterKind::Footer,
                    HeaderFooterType::Default,
                    refs.footer_default.clone().or_else(|| {
                        (!section.title_pg)
                            .then_some(refs.footer_first.clone())
                            .flatten()
                    }),
                ),
                (
                    HeaderFooterKind::Footer,
                    HeaderFooterType::First,
                    section.title_pg.then_some(refs.footer_first).flatten(),
                ),
                (
                    HeaderFooterKind::Footer,
                    HeaderFooterType::Even,
                    refs.footer_even,
                ),
            ];
            for (kind, hf_type, r_id) in descriptors {
                let Some(r_id) = r_id else {
                    continue;
                };
                let mut blocks = self
                    .with_lowered_story(&format!("hf:{r_id}"), render_env, <[LayoutBlock]>::to_vec)
                    .map_err(|error| error.to_string())?;
                for block in &mut blocks {
                    resolve_line_unit_spacing(
                        block,
                        regions.paragraph_spacing_line_px(section_index),
                    );
                    resolve_doc_grid_pitch(block, regions.doc_grid_snap_pitch_px(section_index));
                }
                let metrics = HeaderFooterMetrics {
                    kind,
                    page_size: &page_size,
                    margins: &margins,
                };
                if let Some(variant) = measure_header_footer(
                    r_id,
                    kind,
                    hf_type,
                    section_index,
                    blocks,
                    content_width,
                    metrics,
                    measurement,
                )? {
                    variants.push(variant);
                }
            }
        }
        if variants.is_empty() && regions.watermark.is_none() {
            return Ok(None);
        }
        extend_input_for_header_footer(input, regions, &variants);
        Ok(Some(HeaderFooterPayload {
            title_pg: false,
            even_and_odd_headers: regions.even_and_odd_headers,
            title_page_sections: regions
                .sections
                .iter()
                .enumerate()
                .filter_map(|(index, section)| section.title_pg.then_some(index))
                .collect(),
            even_and_odd_sections: regions
                .sections
                .iter()
                .enumerate()
                .filter_map(|(index, section)| {
                    section
                        .even_and_odd_headers
                        .unwrap_or(regions.even_and_odd_headers)
                        .then_some(index)
                })
                .collect(),
            variants,
            watermark: regions.watermark.clone(),
        }))
    }

    /// Typed resident pagination path shared by the compatibility JSON seam
    /// and `apply_input`.
    fn layout_document_value(&self, input: LayoutInput) -> Result<(), String> {
        let block_fingerprints = measured_fingerprints(&input)?;
        self.layout_document_value_with_fingerprints(input, block_fingerprints)
    }

    /// Paginate a resident measured arena whose clean block fingerprints were
    /// retained while rebuilding the dirty paragraph. Compatibility callers
    /// still enter through `layout_document_value` and fingerprint every block.
    fn layout_document_value_with_fingerprints(
        &self,
        mut input: LayoutInput,
        block_fingerprints: Vec<u64>,
    ) -> Result<(), String> {
        if block_fingerprints.len() != input.measured.len() {
            return Err("resident pagination fingerprints do not match measured blocks".to_owned());
        }
        let input_options_fingerprint = options_fingerprint(&input)?;
        let mut incremental = false;
        let mut deltas = HashMap::new();
        let run = {
            let mut previous = self.pagination.borrow_mut();
            let first_dirty = previous
                .block_fingerprints
                .iter()
                .zip(&block_fingerprints)
                .position(|(previous, next)| previous != next);
            if let Some(dirty_index) = first_dirty
                && incremental_eligible(&previous, &input, input_options_fingerprint)
            {
                let previous = &mut *previous;
                deltas = position_deltas(
                    previous.input.as_ref().expect("eligibility checked input"),
                    &input,
                );
                let attempted = docx_layout::place::layout_document_incremental(
                    &mut input,
                    previous
                        .layout
                        .as_mut()
                        .expect("eligibility checked layout"),
                    &previous.checkpoints,
                    &previous.block_fingerprints,
                    &block_fingerprints,
                    dirty_index,
                );
                match attempted {
                    Ok(run) => {
                        incremental = true;
                        run
                    }
                    Err(docx_layout::LayoutError::Unsupported(_)) => {
                        docx_layout::place::layout_document_checkpointed(&mut input).map_err(
                            |error| match error {
                                docx_layout::LayoutError::Unsupported(_) => {
                                    "UNSUPPORTED".to_owned()
                                }
                                docx_layout::LayoutError::Invalid(reason) => reason,
                            },
                        )?
                    }
                    Err(docx_layout::LayoutError::Invalid(reason)) => return Err(reason),
                }
            } else {
                docx_layout::place::layout_document_checkpointed(&mut input).map_err(|error| {
                    match error {
                        docx_layout::LayoutError::Unsupported(_) => "UNSUPPORTED".to_owned(),
                        docx_layout::LayoutError::Invalid(reason) => reason,
                    }
                })?
            }
        };
        let mut pagination = self.pagination.borrow_mut();
        pagination.input = Some(input);
        pagination.measured_with = None;
        pagination.layout = Some(run.layout);
        pagination.checkpoints = run.checkpoints;
        pagination.block_fingerprints = block_fingerprints;
        pagination.options_fingerprint = input_options_fingerprint;
        pagination.rebuilt_page_start = run.rebuilt_page_start;
        pagination.rebuilt_page_end = run.rebuilt_page_end;
        pagination.position_deltas = deltas;
        pagination.last_incremental = incremental;
        pagination.layout_epoch = pagination.layout_epoch.wrapping_add(1);
        pagination.pagination_calls = pagination.pagination_calls.wrapping_add(1);
        pagination.incremental_pagination_calls = pagination
            .incremental_pagination_calls
            .wrapping_add(u64::from(incremental));
        pagination.pagination_blocks_placed = pagination
            .pagination_blocks_placed
            .wrapping_add(run.placed_blocks as u64);
        Ok(())
    }

    /// Whether the current resident state can complete a plain body-text edit
    /// without consulting the host. This is checked before the document
    /// mutation so `apply_input` cannot discover a missing measurement
    /// template after committing the text.
    pub fn can_apply_input(&self, story: &str, para_id: &str) -> bool {
        if story != "body" {
            return false;
        }
        let render_ready = self.render.borrow().stories.contains_key(story);
        let pagination = self.pagination.borrow();
        let layout_ready = pagination.input.is_some() && pagination.layout.is_some();
        drop(pagination);
        let display_ready = self.display.borrow().extras_json.is_some();
        let measure_ready = self.regions.borrow().is_some()
            || self
                .pagination
                .borrow()
                .input
                .as_ref()
                .and_then(|input| {
                    input.measured.iter().find(|measured| {
                        paragraph_identity(&measured.block)
                            .is_some_and(|(id, _)| block_key(id) == para_id)
                    })
                })
                .is_some_and(|measured| {
                    self.measurement_envelope_for_block(para_id, &measured.block)
                        .is_some()
                });
        render_ready && layout_ready && display_ready && measure_ready
    }

    /// Rebuild the typed measured arena from the newly lowered body story.
    /// Geometry-clean blocks reuse their retained extents while receiving the
    /// new absolute document positions. Only changed paragraph blocks invoke
    /// the resident text measurer.
    fn resident_layout_input(&self, story: &str) -> Result<ResidentLayoutInput, String> {
        self.resident_layout_input_observed(story, &mut || {})
    }

    fn resident_layout_input_observed(
        &self,
        story: &str,
        after_lower: &mut impl FnMut(),
    ) -> Result<ResidentLayoutInput, String> {
        let env = self
            .render
            .borrow()
            .stories
            .get(story)
            .map(|story| story.env.clone())
            .ok_or_else(|| format!("resident render environment missing for story {story:?}"))?;
        self.with_lowered_story_observed(story, &env, after_lower, |blocks| {
            self.resident_layout_input_from_blocks(
                blocks,
                &mut |_, key, previous_block, next_block| {
                    let mut envelope = self
                        .measurement_envelope_for_block(key, previous_block)
                        .ok_or_else(|| {
                            format!("resident measurement template missing for block {key:?}")
                        })?;
                    let fields = envelope.as_object_mut().ok_or_else(|| {
                        "resident measurement envelope is not an object".to_owned()
                    })?;
                    fields.insert(
                        "block".to_owned(),
                        serde_json::to_string(&*next_block)
                            .ok()
                            .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
                            .ok_or_else(|| "serialize dirty paragraph".to_owned())?,
                    );
                    let envelope_json = serde_json::to_string(&envelope)
                        .map_err(|error| format!("serialize measurement envelope: {error}"))?;
                    let extent_json = docx_layout::measure_paragraph_json_resident(&envelope_json)?;
                    let extent: ParagraphExtent = serde_json::from_str(&extent_json)
                        .map_err(|error| format!("parse resident paragraph extent: {error}"))?;
                    Ok(BlockExtent::Paragraph(extent))
                },
            )
        })
        .map_err(|error| error.to_string())?
    }

    /// Shared dirty-block walk over a freshly lowered story: structurally clean
    /// blocks reuse their retained extents (with fresh absolute positions),
    /// changed paragraph blocks are re-measured through `measure_dirty`
    /// (`(block_index, block_key, previous_block, next_block) -> extent`).
    fn resident_layout_input_from_blocks(
        &self,
        blocks: &[LayoutBlock],
        measure_dirty: &mut dyn FnMut(
            usize,
            &str,
            &LayoutBlock,
            &mut LayoutBlock,
        ) -> Result<BlockExtent, String>,
    ) -> Result<ResidentLayoutInput, String> {
        let pagination = self.pagination.borrow();
        let previous = pagination
            .input
            .as_ref()
            .ok_or_else(|| "resident pagination input is not built".to_owned())?;
        let previous_fingerprints = &pagination.block_fingerprints;
        let paragraph_merge = blocks.len().checked_add(1) == Some(previous.measured.len());
        if blocks.len() != previous.measured.len() && !paragraph_merge {
            return Err("resident plain-text input changed the block structure".to_owned());
        }
        if previous.measured.len() != previous_fingerprints.len() {
            return Err("resident pagination fingerprints are not built".to_owned());
        }

        let mut previous_blocks = previous.measured.iter().zip(previous_fingerprints);
        let mut skipped_merged_paragraph = false;
        let mut measured = Vec::with_capacity(blocks.len());
        let mut block_fingerprints = Vec::with_capacity(blocks.len());
        let mut resident_measure_calls = 0_u64;
        let mut resident_reused_blocks = 0_u64;
        for (block_index, next_block) in blocks.iter().enumerate() {
            let mut previous_entry = previous_blocks.next().ok_or_else(|| {
                "resident plain-text input changed the block structure".to_owned()
            })?;
            if paragraph_merge && !resident_block_slots_match(&previous_entry.0.block, next_block) {
                if skipped_merged_paragraph || paragraph_identity(&previous_entry.0.block).is_none()
                {
                    return Err("resident plain-text input changed the block structure".to_owned());
                }
                skipped_merged_paragraph = true;
                previous_entry = previous_blocks.next().ok_or_else(|| {
                    "resident plain-text input changed the block structure".to_owned()
                })?;
            }
            if paragraph_merge && !resident_block_slots_match(&previous_entry.0.block, next_block) {
                return Err("resident plain-text input changed stable block identity".to_owned());
            }
            let (previous_measured, previous_fingerprint) = previous_entry;
            let (Some((next_id, _)), Some((previous_id, _))) = (
                paragraph_identity(next_block),
                paragraph_identity(&previous_measured.block),
            ) else {
                if *next_block != previous_measured.block {
                    return Err(
                        "resident plain-text input changed a non-paragraph block".to_owned()
                    );
                }
                measured.push(MeasuredBlock {
                    block: next_block.clone(),
                    measure: previous_measured.measure.clone(),
                });
                block_fingerprints.push(*previous_fingerprint);
                resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
                continue;
            };
            let key = block_key(next_id);
            if key != block_key(previous_id) {
                return Err("resident plain-text input changed stable block identity".to_owned());
            }
            if *next_block == previous_measured.block {
                measured.push(MeasuredBlock {
                    block: next_block.clone(),
                    measure: previous_measured.measure.clone(),
                });
                block_fingerprints.push(*previous_fingerprint);
                resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
                continue;
            }

            let mut next_measured_block = next_block.clone();
            let measure = measure_dirty(
                block_index,
                &key,
                &previous_measured.block,
                &mut next_measured_block,
            )?;
            let measured_block = MeasuredBlock {
                block: next_measured_block,
                measure,
            };
            block_fingerprints.push(measured_fingerprint(&measured_block)?);
            measured.push(measured_block);
            resident_measure_calls = resident_measure_calls.wrapping_add(1);
        }
        if let Some((removed, _)) = previous_blocks.next() {
            if !paragraph_merge
                || skipped_merged_paragraph
                || paragraph_identity(&removed.block).is_none()
                || previous_blocks.next().is_some()
            {
                return Err("resident plain-text input changed the block structure".to_owned());
            }
            skipped_merged_paragraph = true;
        }
        if paragraph_merge && !skipped_merged_paragraph {
            return Err("resident plain-text input changed the block structure".to_owned());
        }

        let mut measurement = self.measurement.borrow_mut();
        measurement.resident_measure_calls = measurement
            .resident_measure_calls
            .wrapping_add(resident_measure_calls);
        measurement.resident_reused_blocks = measurement
            .resident_reused_blocks
            .wrapping_add(resident_reused_blocks);
        Ok(ResidentLayoutInput {
            input: LayoutInput {
                measured,
                options: previous.options.clone(),
            },
            block_fingerprints,
        })
    }

    /// Reuse retained extents for blocks that cannot have changed (equal
    /// normalized block, width, config, and section-break adjacency).
    /// `Ok(None)` means the caller must measure the whole story.
    fn resident_region_measured(
        &self,
        blocks: &[LayoutBlock],
        widths: &[f64],
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        measurement_fingerprint: u64,
    ) -> Result<Option<(Vec<MeasuredBlock>, Vec<u64>)>, String> {
        let pagination = &mut *self.pagination.borrow_mut();
        let Some(previous) = pagination.input.as_mut() else {
            return Ok(None);
        };
        if pagination.measured_with != Some(measurement_fingerprint)
            || previous.measured.len() != blocks.len()
            || previous.measured.len() != pagination.block_fingerprints.len()
        {
            return Ok(None);
        }
        let previous_widths = region_measurement_widths(
            previous.measured.iter().map(|measured| &measured.block),
            previous,
            regions,
        );
        let previous_fingerprints = &pagination.block_fingerprints;

        let mut measured: Vec<MeasuredBlock> = Vec::with_capacity(blocks.len());
        let mut block_fingerprints = Vec::with_capacity(blocks.len());
        let mut measure_calls = 0_u64;
        let mut reused_blocks = 0_u64;
        let mut section_index = 0_usize;
        for (index, next_block) in blocks.iter().enumerate() {
            // An empty paragraph ahead of a section break measures to a bare
            // mark extent; only its own emptiness plus the next block's kind
            // matter, so a dirty next block has to force a re-measure.
            let next_is_break = matches!(blocks.get(index + 1), Some(LayoutBlock::SectionBreak(_)));
            let retained_next_is_break = matches!(
                previous
                    .measured
                    .get(index + 1)
                    .map(|measured| &measured.block),
                Some(LayoutBlock::SectionBreak(_))
            );
            let previous_entry = &mut previous.measured[index];
            if !resident_block_slots_match(&previous_entry.block, next_block) {
                restore_moved_measures(&mut previous.measured, measured);
                return Ok(None);
            }
            let width_clean = next_is_break == retained_next_is_break
                && widths.get(index) == previous_widths.get(index)
                && matches!(previous_entry.measure, BlockExtent::Unsupported)
                    == matches!(next_block, LayoutBlock::Unsupported);
            if width_clean && *next_block == previous_entry.block {
                measured.push(MeasuredBlock {
                    block: next_block.clone(),
                    measure: std::mem::replace(
                        &mut previous_entry.measure,
                        BlockExtent::Unsupported,
                    ),
                });
                block_fingerprints.push(previous_fingerprints[index]);
                reused_blocks = reused_blocks.wrapping_add(1);
            } else {
                let mut owned = next_block.clone();
                resolve_line_unit_spacing(
                    &mut owned,
                    regions.paragraph_spacing_line_px(section_index),
                );
                resolve_doc_grid_pitch(&mut owned, regions.doc_grid_snap_pitch_px(section_index));
                if let LayoutBlock::SectionBreak(section_break) = &mut owned
                    && let Some(section) = regions.sections.get(section_index)
                {
                    if section.page_size.is_some() {
                        section_break.page_size.clone_from(&section.page_size);
                    }
                    if section.margins.is_some() {
                        section_break.margins.clone_from(&section.margins);
                    }
                    if section.columns.is_some() {
                        section_break.columns.clone_from(&section.columns);
                    }
                }
                suppress_contextual_spacing(blocks, index, &mut owned);
                docx_layout::paragraph_spacing::apply_contextual_spacing_blocks(
                    std::slice::from_mut(&mut owned),
                );
                if width_clean && owned == previous_entry.block {
                    measured.push(MeasuredBlock {
                        block: owned,
                        measure: std::mem::replace(
                            &mut previous_entry.measure,
                            BlockExtent::Unsupported,
                        ),
                    });
                    block_fingerprints.push(previous_fingerprints[index]);
                    reused_blocks = reused_blocks.wrapping_add(1);
                } else {
                    let measure = if next_is_break
                        && matches!(&owned, LayoutBlock::Paragraph(paragraph) if paragraph.runs.is_empty())
                    {
                        BlockExtent::Paragraph(ParagraphExtent {
                            lines: Vec::new(),
                            total_height: 0.0,
                        })
                    } else {
                        match docx_layout::measure_blocks::measure_block(
                            &mut owned,
                            widths.get(index).copied().unwrap_or(0.0),
                            measurement,
                        ) {
                            Ok(measure) => measure,
                            Err(error) => {
                                restore_moved_measures(&mut previous.measured, measured);
                                return Err(error);
                            }
                        }
                    };
                    let entry = MeasuredBlock {
                        block: owned,
                        measure,
                    };
                    let fingerprint = match measured_fingerprint(&entry) {
                        Ok(fingerprint) => fingerprint,
                        Err(error) => {
                            restore_moved_measures(&mut previous.measured, measured);
                            return Err(error);
                        }
                    };
                    block_fingerprints.push(fingerprint);
                    measured.push(entry);
                    measure_calls = measure_calls.wrapping_add(1);
                }
            }
            if matches!(next_block, LayoutBlock::SectionBreak(_)) {
                section_index += 1;
            }
        }
        let mut measurement_state = self.measurement.borrow_mut();
        measurement_state.resident_measure_calls = measurement_state
            .resident_measure_calls
            .wrapping_add(measure_calls);
        measurement_state.resident_reused_blocks = measurement_state
            .resident_reused_blocks
            .wrapping_add(reused_blocks);
        Ok(Some((measured, block_fingerprints)))
    }

    /// Complete the post-edit dependency cone and return its binary frame.
    /// No measured/layout/display values cross the wasm boundary.
    pub fn apply_and_layout(
        &self,
        story: &str,
        expected_frame_epoch: u64,
    ) -> Result<Vec<u8>, String> {
        if self.regions.borrow().is_some() {
            if !self.apply_and_layout_regions_resident(story, &mut |_| {})? {
                self.apply_and_layout_regions_full()?;
            }
            let extras = self.resident_region_display_extras()?;
            return self.build_display_list_frame(&extras, expected_frame_epoch);
        }
        let resident = self.resident_layout_input(story)?;
        self.layout_document_value_with_fingerprints(resident.input, resident.block_fingerprints)?;
        let extras = self
            .display
            .borrow()
            .extras_json
            .clone()
            .ok_or_else(|| "resident display extras are not built".to_owned())?;
        self.build_display_list_frame(&extras, expected_frame_epoch)
    }

    /// Fallback for edits the resident region path cannot absorb: replay the
    /// retained region request through the full pass (no serialization).
    fn apply_and_layout_regions_full(&self) -> Result<(), String> {
        let request = self
            .regions
            .borrow()
            .as_ref()
            .map(|state| state.request_json.clone())
            .expect("region state checked by the caller");
        self.layout_document_with_regions_value(&request)?;
        Ok(())
    }

    /// Absorb a plain body-text edit into retained region state: re-lower the
    /// body, re-measure only fingerprint-dirty paragraphs at their retained
    /// section widths, paginate (incrementally when eligible), and re-stamp
    /// page regions. `Ok(false)` means the caller must run the full pass —
    /// float-anchoring content, notes, a structural change, or a page-count
    /// change (header/footer PAGE/NUMPAGES field widths may depend on it).
    fn apply_and_layout_regions_resident(
        &self,
        story: &str,
        phase: &mut impl FnMut(RegionResidentPhase),
    ) -> Result<bool, String> {
        if story != "body" {
            return Ok(false);
        }
        let fast_config = {
            let state = self.regions.borrow();
            state.as_ref().and_then(|state| {
                let fast = state.fast_path.as_ref()?;
                fast.notes_clear.then(|| {
                    (
                        Rc::clone(&fast.regions),
                        Rc::clone(&fast.measurement),
                        fast.measurement_fingerprint,
                    )
                })
            })
        };
        let Some((regions, measurement, measurement_fingerprint)) = fast_config else {
            return Ok(false);
        };
        let env = {
            let render = self.render.borrow();
            let Some(lowered) = render.stories.get(story) else {
                return Ok(false);
            };
            lowered.env.clone()
        };
        let outcome = self
            .with_lowered_story_observed(
                story,
                &env,
                &mut || phase(RegionResidentPhase::Lowered),
                |blocks| -> Result<Option<(ResidentLayoutInput, usize)>, String> {
                    let (widths, geometry, previous_pages) = {
                        let pagination = self.pagination.borrow();
                        let (Some(input), Some(layout)) =
                            (pagination.input.as_ref(), pagination.layout.as_ref())
                        else {
                            return Ok(None);
                        };
                        (
                            region_measurement_widths(blocks.iter(), input, &regions),
                            initial_float_page_geometry(input, &regions),
                            layout.pages.len(),
                        )
                    };
                    let default_width = widths.first().copied().unwrap_or(0.0);
                    if docx_layout::measure_blocks::has_floating_zones(
                        blocks,
                        default_width,
                        measurement.as_ref(),
                        Some(&geometry),
                    )? {
                        return Ok(None);
                    }
                    match self.resident_layout_input_from_blocks(
                        blocks,
                        &mut |index, _key, _previous_block, next_block| {
                            let width = widths.get(index).copied().unwrap_or(default_width);
                            resolve_line_unit_spacing(
                                next_block,
                                regions.paragraph_spacing_line_px(0),
                            );
                            resolve_doc_grid_pitch(next_block, regions.doc_grid_snap_pitch_px(0));
                            docx_layout::measure_blocks::measure_block(
                                next_block,
                                width,
                                measurement.as_ref(),
                            )
                        },
                    ) {
                        Ok(resident) => Ok(Some((resident, previous_pages))),
                        Err(_) => Ok(None),
                    }
                },
            )
            .map_err(|error| error.to_string())??;
        let (resident, previous_pages) = match outcome {
            Some(resident) => resident,
            None => return Ok(false),
        };
        phase(RegionResidentPhase::Measured);
        self.layout_document_value_with_fingerprints(resident.input, resident.block_fingerprints)?;
        let mut pagination = self.pagination.borrow_mut();
        // The fast path measures through the region config too, so its
        // retained arena is also eligible for the next pass's reuse walk.
        pagination.measured_with = Some(measurement_fingerprint);
        let layout = pagination
            .layout
            .as_mut()
            .expect("layout retained after successful pagination");
        apply_document_regions(layout, &regions);
        Ok(layout.pages.len() == previous_pages)
    }

    fn resident_region_display_extras(&self) -> Result<String, String> {
        let extras = self
            .display
            .borrow()
            .extras_json
            .clone()
            .ok_or_else(|| "resident display extras are not built".to_owned())?;
        let mut value: serde_json::Value = serde_json::from_str(&extras)
            .map_err(|error| format!("parse display extras: {error}"))?;
        let fields = value
            .as_object_mut()
            .ok_or_else(|| "resident display extras must be an object".to_owned())?;
        let headers_footers = self
            .regions
            .borrow()
            .as_ref()
            .and_then(|state| state.headers_footers.clone());
        if let Some(headers_footers) = headers_footers {
            fields.insert("headersFooters".to_owned(), headers_footers);
        } else {
            fields.remove("headersFooters");
        }
        serde_json::to_string(&value).map_err(|error| format!("serialize display extras: {error}"))
    }

    /// Profiled twin of [`Self::apply_and_layout`]. The caller supplies a
    /// monotonic millisecond clock (the worker's `performance.now`) so this
    /// module stays browser-agnostic and the production method above remains
    /// timer-free.
    ///
    /// Attribution: on the region fast path, lower/measure/paginate are split
    /// exactly like the plain resident path. When the fast path falls back,
    /// the full region pass (its lowering, measurement, notes, and
    /// pagination) lands in `paginate_ms` as one lump — the same attribution
    /// the region path had before stage splitting existed.
    pub fn apply_and_layout_profiled(
        &self,
        story: &str,
        expected_frame_epoch: u64,
        now: &mut impl FnMut() -> f64,
    ) -> Result<(Vec<u8>, EngineApplyProfile), String> {
        let mut profile = EngineApplyProfile::default();
        let mut started = now();
        let has_regions = self.regions.borrow().is_some();
        let extras;
        if has_regions {
            let fast = self.apply_and_layout_regions_resident(story, &mut |mark| {
                let finished = now();
                match mark {
                    RegionResidentPhase::Lowered => profile.lower_ms = finished - started,
                    RegionResidentPhase::Measured => profile.measure_ms = finished - started,
                }
                started = finished;
            })?;
            if !fast {
                self.apply_and_layout_regions_full()?;
            }
            let finished = now();
            profile.paginate_ms = finished - started;
            started = finished;
            extras = self.resident_region_display_extras()?;
        } else {
            let resident = self.resident_layout_input_observed(story, &mut || {
                let finished = now();
                profile.lower_ms = finished - started;
                started = finished;
            })?;
            let finished = now();
            profile.measure_ms = finished - started;
            started = finished;

            self.layout_document_value_with_fingerprints(
                resident.input,
                resident.block_fingerprints,
            )?;
            let finished = now();
            profile.paginate_ms = finished - started;
            started = finished;

            extras = self
                .display
                .borrow()
                .extras_json
                .clone()
                .ok_or_else(|| "resident display extras are not built".to_owned())?;
        }
        let mut display_phase = 0;
        let bytes =
            self.build_display_list_frame_observed(&extras, expected_frame_epoch, &mut || {
                let finished = now();
                if display_phase == 0 {
                    profile.display_input_ms = finished - started;
                } else if display_phase == 1 {
                    profile.display_build_ms = finished - started;
                } else {
                    profile.display_finalize_ms = finished - started;
                    profile.display_ms = profile.display_input_ms
                        + profile.display_build_ms
                        + profile.display_finalize_ms;
                }
                started = finished;
                display_phase += 1;
            })?;
        profile.encode_ms = now() - started;
        Ok((bytes, profile))
    }

    /// Builds and retains the typed display list.
    pub fn build_display_list_json(&self, input_json: &str) -> Result<String, String> {
        let list = docx_layout::build_display_list_value(input_json)?;
        let display_json =
            serde_json::to_string(&list).map_err(|error| format!("serialize: {error}"))?;
        let mut display = self.display.borrow_mut();
        display.list = Some(list);
        display.resident_input = None;
        display.frame_epoch = display.frame_epoch.wrapping_add(1);
        display.display_builds = display.display_builds.wrapping_add(1);
        Ok(display_json)
    }

    /// Build the retained display list and return a binary FrameDelta v1.
    /// `expected_frame_epoch` is the last frame the host actually applied. A
    /// mismatch automatically widens to a full recovery frame.
    pub fn build_display_list_frame(
        &self,
        extras_json: &str,
        expected_frame_epoch: u64,
    ) -> Result<Vec<u8>, String> {
        self.build_display_list_frame_observed(extras_json, expected_frame_epoch, &mut || {})
    }

    fn build_display_list_frame_observed(
        &self,
        extras_json: &str,
        expected_frame_epoch: u64,
        observe_display_phase: &mut impl FnMut(),
    ) -> Result<Vec<u8>, String> {
        let extras_fingerprint = hash_bytes(extras_json.as_bytes());
        let (incremental_build, rebuilt_display_pages, rebuilt_page_start, rebuilt_page_end) = {
            let pagination = self.pagination.borrow();
            let input = pagination
                .input
                .as_ref()
                .ok_or_else(|| "resident pagination input is not built".to_owned())?;
            let layout = pagination
                .layout
                .as_ref()
                .ok_or_else(|| "resident layout is not built".to_owned())?;
            let mut display = self.display.borrow_mut();
            if pagination.last_incremental && display.extras_fingerprint == extras_fingerprint {
                let rebuilt_pages = pagination
                    .rebuilt_page_end
                    .saturating_sub(pagination.rebuilt_page_start);
                let incremental = if let DisplayState {
                    list: Some(previous),
                    resident_input: Some(resident_input),
                    ..
                } = &mut *display
                {
                    docx_layout::update_resident_display_list_incremental_observed(
                        input,
                        layout,
                        resident_input,
                        previous,
                        pagination.rebuilt_page_start,
                        pagination.rebuilt_page_end,
                        &pagination.position_deltas,
                        observe_display_phase,
                    )?
                } else {
                    false
                };
                if !incremental {
                    let (resident_input, list) = docx_layout::build_resident_display_list_observed(
                        input,
                        layout,
                        extras_json,
                        observe_display_phase,
                    )?;
                    display.resident_input = Some(resident_input);
                    display.list = Some(list);
                }
                (
                    incremental,
                    rebuilt_pages,
                    pagination.rebuilt_page_start,
                    pagination.rebuilt_page_end,
                )
            } else {
                let (resident_input, list) = docx_layout::build_resident_display_list_observed(
                    input,
                    layout,
                    extras_json,
                    observe_display_phase,
                )?;
                display.resident_input = Some(resident_input);
                display.list = Some(list);
                (false, layout.pages.len(), 0, layout.pages.len())
            }
        };
        observe_display_phase();
        let mut display = self.display.borrow_mut();
        display.frame_epoch = display.frame_epoch.wrapping_add(1);
        display.display_builds = display.display_builds.wrapping_add(1);
        display.incremental_display_builds = display
            .incremental_display_builds
            .wrapping_add(u64::from(incremental_build));
        display.rebuilt_display_pages = display
            .rebuilt_display_pages
            .wrapping_add(rebuilt_display_pages as u64);
        display.extras_fingerprint = extras_fingerprint;
        display.extras_json = Some(extras_json.to_owned());
        let frame_epoch = display.frame_epoch;
        let binary_frame_epoch = display.binary_frame_epoch;
        let full = expected_frame_epoch != binary_frame_epoch || binary_frame_epoch == 0;
        let layout_epoch = self.pagination.borrow().layout_epoch;
        let mut next_page_id = display.next_page_id;
        // Split borrows: the encoder reads the retained list and the previous
        // snapshots in place — no per-frame deep clone of the snapshot set.
        let display = &mut *display;
        let previous_pages = &display.pages;
        let list = display
            .list
            .as_ref()
            .expect("display list built before FrameDelta encoding");
        let epochs = FrameEpochs {
            doc_epoch: self.doc_epoch(),
            layout_epoch,
            frame_epoch,
            base_frame_epoch: binary_frame_epoch,
        };
        let (bytes, pages) =
            if incremental_build && !full && previous_pages.len() == list.pages.len() {
                encode_frame_delta_incremental(
                    list,
                    previous_pages,
                    epochs,
                    &mut next_page_id,
                    rebuilt_page_start..rebuilt_page_end,
                )?
            } else {
                encode_frame_delta(list, previous_pages, epochs, full, &mut next_page_id)?
            };
        display.pages = pages;
        display.next_page_id = next_page_id;
        display.binary_frame_epoch = frame_epoch;
        Ok(bytes)
    }

    /// Read the resident display list without cloning or serializing it.
    #[allow(dead_code)]
    pub fn with_display_list<T>(&self, read: impl FnOnce(&DisplayList) -> T) -> Option<T> {
        self.display.borrow().list.as_ref().map(read)
    }

    pub fn resident_caret_snapshot(
        &self,
        paragraph: Option<(&str, u32)>,
    ) -> Result<ResidentCaretSnapshot, String> {
        let position = paragraph.and_then(|(para_id, offset)| {
            self.pagination
                .borrow()
                .input
                .as_ref()?
                .measured
                .iter()
                .find_map(|measured| {
                    let (id, start) = paragraph_identity(&measured.block)?;
                    (block_key(id) == para_id).then_some(start?)
                })
                .and_then(|start| {
                    (start.is_finite()
                        && start.fract() == 0.0
                        && start >= i64::MIN as f64
                        && start <= i64::MAX as f64)
                        .then_some(start as i64)
                })
                .and_then(|start| start.checked_add(1 + i64::from(offset)))
        });
        let display = self.display.borrow();
        let frame_epoch = display.binary_frame_epoch;
        let caret_rect = position
            .and_then(|position| {
                display
                    .list
                    .as_ref()
                    .and_then(|list| docx_layout::hit::caret_rect(list, position))
            })
            .and_then(|rect: CaretRect| {
                let page = display.pages.get(rect.page_index)?;
                Some(ResidentCaretRect {
                    page_index: rect.page_index,
                    page_id: page.page_id.to_string(),
                    x: rect.x,
                    y: rect.y,
                    height: rect.height,
                })
            });
        Ok(ResidentCaretSnapshot {
            frame_epoch,
            caret_rect,
        })
    }

    /// Region-aware hit testing directly against the resident display list.
    pub fn display_hit_test_regions_json(
        &self,
        page_index: usize,
        x: f64,
        y: f64,
    ) -> Result<String, String> {
        self.with_display_list(|list| docx_layout::hit::hit_test_regions(list, page_index, x, y))
            .ok_or_else(|| "resident display list is not built".to_owned())
            .and_then(|hit| serde_json::to_string(&hit).map_err(|error| error.to_string()))
    }

    pub fn display_vertical_move_json(
        &self,
        position: i64,
        direction: &str,
        goal_x: f64,
    ) -> Result<String, String> {
        let direction = match direction {
            "up" => VerticalDirection::Up,
            "down" => VerticalDirection::Down,
            other => return Err(format!("unknown vertical direction {other:?}")),
        };
        self.with_display_list(|list| {
            docx_layout::hit::vertical_move(
                list,
                position,
                direction,
                goal_x.is_finite().then_some(goal_x),
            )
        })
        .ok_or_else(|| "resident display list is not built".to_owned())
        .and_then(|movement| serde_json::to_string(&movement).map_err(|error| error.to_string()))
    }

    /// Body range geometry directly against the resident display list.
    pub fn display_range_rects_json(&self, from: i64, to: i64) -> Result<String, String> {
        self.with_display_list(|list| docx_layout::hit::range_rects(list, from, to))
            .ok_or_else(|| "resident display list is not built".to_owned())
            .and_then(|rects| serde_json::to_string(&rects).map_err(|error| error.to_string()))
    }

    /// Body, header/footer and note range geometry directly against the
    /// resident list.
    pub fn display_range_rects_region_json(
        &self,
        region: &str,
        part_id: &str,
        from: i64,
        to: i64,
    ) -> Result<String, String> {
        let scope = docx_layout::hit::parse_region_scope(region, part_id)?;
        self.with_display_list(|list| {
            docx_layout::hit::range_rects_in_region(list, scope, from, to)
        })
        .ok_or_else(|| "resident display list is not built".to_owned())
        .and_then(|rects| serde_json::to_string(&rects).map_err(|error| error.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use yrs::Any;
    use yrs::types::Attrs;

    #[test]
    fn lowered_story_is_resident_and_generation_tagged() {
        let engine = EngineSession::new(7);
        engine
            .doc()
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        let env = RenderEnv::default();

        let first = engine.lower_story_json("body", &env).unwrap();
        let expected = serde_json::to_string(
            &crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap(),
        )
        .unwrap();
        assert_eq!(
            first, expected,
            "resident output must match the uncached lowering"
        );
        let second = engine.lower_story_json("body", &env).unwrap();
        assert_eq!(first, second);
        assert_eq!(
            engine.stats(),
            EngineStats {
                doc_epoch: 1,
                lowered_story_count: 1,
                lowered_block_count: 1,
                lower_cache_hits: 1,
                lower_cache_misses: 1,
                retained_measure_templates: 0,
                compatibility_measure_calls: 0,
                resident_measure_calls: 0,
                resident_reused_blocks: 0,
                layout_epoch: 0,
                retained_measured_blocks: 0,
                retained_pages: 0,
                pagination_calls: 0,
                incremental_pagination_calls: 0,
                pagination_blocks_placed: 0,
                retained_checkpoints: 0,
                rebuilt_pages: 0,
                frame_epoch: 0,
                retained_display_pages: 0,
                retained_display_primitives: 0,
                display_builds: 0,
                incremental_display_builds: 0,
                rebuilt_display_pages: 0,
            }
        );

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                "!",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let third = engine.lower_story_json("body", &env).unwrap();
        assert_ne!(first, third);
        assert_eq!(engine.stats().doc_epoch, 2);
        assert_eq!(engine.stats().lower_cache_misses, 2);
    }

    #[test]
    fn render_environment_participates_in_cache_identity() {
        let engine = EngineSession::new(11);
        engine
            .doc()
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        let original = RenderEnv::default();
        engine.lower_story_json("body", &original).unwrap();

        let mut changed = original.clone();
        changed.default_tab_stop_twips = Some(720.0);
        engine.lower_story_json("body", &changed).unwrap();

        assert_eq!(engine.stats().lower_cache_hits, 0);
        assert_eq!(engine.stats().lower_cache_misses, 2);
    }

    #[test]
    fn pagination_input_and_layout_are_retained_with_json() {
        let engine = EngineSession::new(13);
        let input = r#"{
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            }
        }"#;
        let resident = engine.layout_document_json(input).unwrap();
        let expected = docx_layout::layout_to_json(input).unwrap();
        assert_eq!(resident, expected);
        assert_eq!(engine.stats().layout_epoch, 1);
        assert_eq!(engine.stats().retained_measured_blocks, 0);
        assert_eq!(engine.stats().retained_pages, 1);
        assert_eq!(engine.stats().pagination_calls, 1);
    }

    #[test]
    fn region_layout_operation_stamps_pages_and_returns_render_envelope() {
        let engine = EngineSession::new(131);
        let request = serde_json::json!({
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            },
            "regions": {
                "sections": [{
                    "sectionId": "main",
                    "pageNumbering": {"start": 7, "format": "upperRoman"},
                    "headerDistance": 24,
                    "headerFooterRefs": {"headerDefault": "rId7"}
                }],
                "headersFooters": {"variants": []}
            }
        });

        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();

        assert_eq!(output["measured"], serde_json::json!([]));
        assert_eq!(output["layout"]["pages"][0]["sectionId"], "main");
        assert_eq!(output["layout"]["pages"][0]["sectionIndex"], 0);
        assert_eq!(output["layout"]["pages"][0]["sectionPageIndex"], 0);
        assert_eq!(output["layout"]["pages"][0]["sectionPageNumber"], 7);
        assert_eq!(output["layout"]["pages"][0]["pageLabel"], "VII");
        assert_eq!(
            output["layout"]["pages"][0]["headerDistance"].as_f64(),
            Some(24.0)
        );
        assert_eq!(
            output["layout"]["pages"][0]["headerFooterRefs"]["headerDefault"],
            "rId7"
        );
        assert_eq!(
            output["headersFooters"],
            serde_json::json!({"variants": []})
        );
        assert_eq!(engine.stats().retained_pages, 1);
    }

    #[test]
    fn region_layout_accepts_a_next_column_section_start() {
        let engine = EngineSession::new(132);
        let request = serde_json::json!({
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            },
            "regions": {
                "sections": [{
                    "sectionId": "main",
                    "properties": {"sectionStart": "nextColumn", "columnCount": 2}
                }]
            }
        });

        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        assert_eq!(output["options"]["bodyBreakType"], "nextColumn");
    }

    /// One 200x120 page (10px margins, so a content box of x 10..190 /
    /// y 10..110) whose single paragraph carries a footnote reference.
    fn note_area_layout_request() -> serde_json::Value {
        serde_json::json!({
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "p1",
                    "runs": [{
                        "kind": "text",
                        "text": "x",
                        "pmStart": 1,
                        "pmEnd": 2,
                        "footnoteRefId": 7
                    }],
                    "pmStart": 0,
                    "pmEnd": 2
                },
                "measure": {
                    "kind": "paragraph",
                    "lines": [{
                        "headRun": 0,
                        "headChar": 0,
                        "tailRun": 0,
                        "tailChar": 1,
                        "width": 10,
                        "ascent": 8,
                        "descent": 2,
                        "lineHeight": 20
                    }],
                    "totalHeight": 20
                }
            }],
            "options": {
                "pageSize": {"w": 200, "h": 120},
                "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}
            },
            "regions": {
                "noteSettings": {
                    "footnote": {"numFmt": "upperRoman", "numStart": 3}
                },
                "sections": [{
                    "sectionId": "main",
                    "noteSettings": {"footnoteColumns": 2}
                }]
            },
            "notes": {
                "contents": [{
                    "id": 7,
                    "height": 20,
                    "blocks": [{
                        "kind": "paragraph",
                        "id": "note-p",
                        "runs": [{"kind": "text", "text": "note", "pmStart": 1, "pmEnd": 5}],
                        "pmStart": 1,
                        "pmEnd": 6
                    }],
                    "measures": [{
                        "kind": "paragraph",
                        "lines": [{
                            "headRun": 0,
                            "headChar": 0,
                            "tailRun": 0,
                            "tailChar": 4,
                            "width": 40,
                            "ascent": 8,
                            "descent": 2,
                            "lineHeight": 20
                        }],
                        "totalHeight": 20
                    }]
                }]
            }
        })
    }

    #[test]
    fn region_layout_operation_stabilizes_and_emits_note_areas() {
        let engine = EngineSession::new(132);
        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&note_area_layout_request().to_string())
                .unwrap(),
        )
        .unwrap();
        let page = &output["layout"]["pages"][0];

        assert_eq!(output["notesConverged"], true);
        assert_eq!(page["footnoteIds"], serde_json::json!([7.0]));
        assert_eq!(page["footnoteReservedHeight"], 32.0);
        assert_eq!(page["footnoteColumns"], 2.0);
        assert_eq!(page["noteAreas"][0]["kind"], "footnote");
        assert_eq!(page["noteAreas"][0]["columns"], 2);
        assert_eq!(page["noteAreas"][0]["notes"][0]["displayLabel"], "III");
        assert_eq!(engine.stats().pagination_calls, 2);
    }

    /// A laid-out note is reachable through the resident hit test: the point
    /// addresses the note's own story and its glyphs invite typing.
    #[test]
    fn a_laid_out_note_area_resolves_to_its_own_story() {
        let engine = EngineSession::new(133);
        let output = engine
            .layout_document_with_regions_json(&note_area_layout_request().to_string())
            .unwrap();
        engine.build_display_list_json(&output).unwrap();
        let run = engine
            .with_display_list(|list| {
                let run = list.pages[0].note_areas[0]
                    .primitives
                    .iter()
                    .find_map(|primitive| match primitive {
                        docx_layout::display_list::Primitive::Text(run) => Some(run),
                        _ => None,
                    })
                    .expect("the note area paints its text");
                (
                    run.x.as_f64().unwrap() + run.width.as_f64().unwrap() / 2.0,
                    run.baseline_y.as_f64().unwrap(),
                )
            })
            .expect("the display list is built");

        let hit = |x: f64, y: f64| -> serde_json::Value {
            serde_json::from_str(&engine.display_hit_test_regions_json(0, x, y).unwrap()).unwrap()
        };
        let note = hit(run.0, run.1 - 2.0);
        assert_eq!(note["region"], "footnote");
        assert_eq!(note["noteId"], 7);
        assert_eq!(note["target"], "text");
        assert!(note["pos"].is_i64(), "the note resolved no position");
        // the body line above the area stays the body's
        assert_eq!(hit(20.0, 20.0)["region"], "body");
    }

    #[test]
    fn region_layout_operation_measures_header_story_and_extends_margin() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(133);
        engine
            .doc()
            .create_story("hf:rId1", "Header", "Normal", "left")
            .unwrap();
        let request = serde_json::json!({
            "measured": [],
            "options": {
                "pageSize": {"w": 300, "h": 200},
                "margins": {
                    "top": 20,
                    "right": 20,
                    "bottom": 20,
                    "left": 20,
                    "header": 5,
                    "footer": 5
                }
            },
            "regions": {
                "sections": [{
                    "sectionId": "main",
                    "pageSize": {"w": 300, "h": 200},
                    "margins": {
                        "top": 20,
                        "right": 20,
                        "bottom": 20,
                        "left": 20,
                        "header": 5,
                        "footer": 5
                    },
                    "headerFooterRefs": {"headerFirst": "rId1"}
                }]
            },
            "measurement": {
                "fontChains": {"liberation sans|0|0": [font_id]},
                "defaults": {"fontSize": 24, "fontFamily": "Liberation Sans"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });

        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let value: serde_json::Value = serde_json::from_str(&output).unwrap();

        assert_eq!(value["headersFooters"]["variants"][0]["rId"], "rId1");
        assert_eq!(value["headersFooters"]["variants"][0]["type"], "default");
        assert_eq!(
            value["headersFooters"]["variants"][0]["measured"][0]["block"]["kind"],
            "paragraph"
        );
        assert!(value["options"]["margins"]["top"].as_f64().unwrap() > 20.0);
        assert_eq!(
            value["layout"]["pages"][0]["headerFooterRefs"]["headerFirst"],
            "rId1"
        );

        let display: serde_json::Value =
            serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
        assert_eq!(display["pages"][0]["header"]["rId"], "rId1");
        assert!(
            !display["pages"][0]["header"]["primitives"]
                .as_array()
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn region_layout_operation_lowers_and_measures_resident_body() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(134);
        engine
            .doc()
            .create_story("body", "Resident body", "Normal", "left")
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "options": {"pageGap": 20},
            "regions": {
                "sections": [{
                    "sectionId": "main",
                    "pageSize": {"w": 300, "h": 200},
                    "margins": {"top": 20, "right": 20, "bottom": 20, "left": 20}
                }]
            },
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });

        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();

        assert_eq!(output["measured"][0]["block"]["kind"], "paragraph");
        assert_eq!(
            output["measured"][0]["block"]["runs"][0]["text"],
            "Resident body"
        );
        assert!(
            output["measured"][0]["measure"]["totalHeight"]
                .as_f64()
                .unwrap()
                > 0.0
        );
        assert_eq!(output["layout"]["pages"][0]["sectionId"], "main");
        assert_eq!(engine.stats().retained_measured_blocks, 1);
    }

    #[test]
    fn resident_region_layout_reflows_after_input_without_host_measurement() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(137);
        engine
            .doc()
            .create_story("body", "Before", "Normal", "left")
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "sectionId": "main",
                "properties": {
                    "pageWidth": 4320,
                    "pageHeight": 2880,
                    "marginTop": 300,
                    "marginRight": 300,
                    "marginBottom": 300,
                    "marginLeft": 300
                }
            }]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        engine
            .build_display_list_frame(
                &serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0,
            )
            .unwrap();
        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        assert!(engine.can_apply_input("body", &paragraph.para_id));
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 6),
                " after",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();

        let frame = engine.apply_and_layout("body", 1).unwrap();
        let pagination = engine.pagination.borrow();
        let input = pagination.input.as_ref().unwrap();
        let LayoutBlock::Paragraph(paragraph) = &input.measured[0].block else {
            panic!("paragraph expected");
        };
        let Run::Text(text) = &paragraph.runs[0] else {
            panic!("text expected");
        };

        assert_eq!(text.text, "Before after");
        assert_eq!(
            pagination.layout.as_ref().unwrap().pages[0]
                .section_id
                .as_deref(),
            Some("main")
        );
        assert!(!frame.is_empty());
    }

    #[test]
    fn resident_region_fast_path_reuses_clean_blocks_and_matches_the_full_pass() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(138);
        engine
            .doc()
            .create_story("body", "AlphaBravo", "Normal", "left")
            .unwrap();
        engine
            .doc()
            .split_paragraph(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                None,
            )
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "sectionId": "main",
                "properties": {
                    "pageWidth": 4320,
                    "pageHeight": 2880,
                    "marginTop": 300,
                    "marginRight": 300,
                    "marginBottom": 300,
                    "marginLeft": 300
                }
            }]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        engine
            .build_display_list_frame(
                &serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0,
            )
            .unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 2),
                "xx",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();

        let stats_before = engine.stats();
        engine.apply_and_layout("body", 1).unwrap();
        let stats_after = engine.stats();
        assert_eq!(
            stats_after.resident_measure_calls,
            stats_before.resident_measure_calls + 1,
            "only the dirty paragraph re-measures on the region fast path"
        );
        assert_eq!(
            stats_after.resident_reused_blocks,
            stats_before.resident_reused_blocks + 1,
            "the clean paragraph reuses its retained extent"
        );

        let fast_json = {
            let pagination = engine.pagination.borrow();
            let regions_state = engine.regions.borrow();
            serialize_region_layout(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                regions_state.as_ref().unwrap().headers_footers.as_ref(),
                true,
            )
            .unwrap()
        };
        let full_json = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        assert_eq!(
            fast_json, full_json,
            "resident region fast path state is byte-identical to a full pass"
        );
    }

    /// Compares resident and full region passes on a paragraph-heavy document.
    /// Run: `cargo test -p betteroffice-docx-edit --release --lib -- --ignored perf_probe --nocapture`
    #[test]
    #[ignore = "perf probe; run explicitly with --ignored --nocapture"]
    fn perf_probe_region_apply_input() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        const PARAGRAPHS: usize = 200;
        const EDITS: u32 = 30;
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(139);
        engine
            .doc()
            .create_story("body", "", "Normal", "left")
            .unwrap();
        let ctx = crate::EditCtx::local("", "");
        let mut cursor = 0_u32;
        for index in 0..PARAGRAPHS {
            let text = format!("Paragraph {index}: the quick brown fox jumps over the lazy dog.");
            engine
                .doc()
                .insert_text(
                    &ctx,
                    crate::Position::new("body", cursor),
                    &text,
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            cursor += text.chars().count() as u32;
            if index + 1 < PARAGRAPHS {
                engine
                    .doc()
                    .split_paragraph(&ctx, crate::Position::new("body", cursor), None)
                    .unwrap();
                cursor += 1;
            }
        }
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let extras = serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string();
        let started = std::time::Instant::now();
        engine.layout_document_with_regions_json(&request).unwrap();
        let cold_layout = started.elapsed();
        engine.build_display_list_frame(&extras, 0).unwrap();
        println!("cold first layout over {PARAGRAPHS} paragraphs: {cold_layout:?}");

        let mut epoch = engine.display.borrow().binary_frame_epoch;
        let started = std::time::Instant::now();
        for edit in 0..EDITS {
            engine
                .doc()
                .insert_text(
                    &ctx,
                    crate::Position::new("body", 12 + edit),
                    "x",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            engine.apply_and_layout("body", epoch).unwrap();
            epoch = engine.display.borrow().binary_frame_epoch;
        }
        let fast = started.elapsed();

        let started = std::time::Instant::now();
        for edit in 0..EDITS {
            engine
                .doc()
                .insert_text(
                    &ctx,
                    crate::Position::new("body", 42 + edit),
                    "x",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            // Full region pass with serialized envelope.
            engine.layout_document_with_regions_json(&request).unwrap();
            let merged = engine.resident_region_display_extras().unwrap();
            engine.build_display_list_frame(&merged, epoch).unwrap();
            epoch = engine.display.borrow().binary_frame_epoch;
        }
        let full = started.elapsed();

        let mut clock = || {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_secs_f64()
                * 1000.0
        };
        engine
            .doc()
            .insert_text(
                &ctx,
                crate::Position::new("body", 7),
                "y",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let (_, profile) = engine
            .apply_and_layout_profiled("body", epoch, &mut clock)
            .unwrap();
        let stats = engine.stats();
        println!(
            "region apply_input over {PARAGRAPHS} paragraphs x {EDITS} edits: \
             fast path {:?}/edit, full pass {:?}/edit ({:.1}x); \
             incremental paginations {}, incremental display builds {}, \
             resident measures {}, reused blocks {}; profile {profile:?}",
            fast / EDITS,
            full / EDITS,
            full.as_secs_f64() / fast.as_secs_f64(),
            stats.incremental_pagination_calls,
            stats.incremental_display_builds,
            stats.resident_measure_calls,
            stats.resident_reused_blocks,
        );
    }

    #[test]
    fn region_font_preflight_returns_requirements_without_layout_blocks() {
        let engine = EngineSession::new(136);
        engine
            .doc()
            .create_story("body", "Resident body", "Normal", "left")
            .unwrap();
        engine
            .doc()
            .create_story("hf:rId1", "Header", "Normal", "left")
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "headerFooterRefs": {"headerDefault": "rId1"}
            }]},
            "renderEnv": {}
        });

        let requirements: serde_json::Value = serde_json::from_str(
            &engine
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();

        assert_eq!(requirements[0]["key"], "calibri|0|0");
        assert!(requirements[0].get("blocks").is_none());
        assert_eq!(engine.stats().layout_epoch, 0);
        assert_eq!(engine.stats().retained_measured_blocks, 0);
    }

    #[test]
    fn region_layout_operation_lowers_measures_and_places_resident_note_story() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(135);
        engine
            .doc()
            .create_story("body", "", "Normal", "left")
            .unwrap();
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![
                    crate::RawOp::Delete { index: 0, len: 1 },
                    crate::RawOp::Insert {
                        index: 0,
                        text: "See ".to_owned(),
                        attrs: Attrs::new(),
                    },
                    crate::RawOp::InsertEmbed {
                        index: 4,
                        kind: "noteRef".to_owned(),
                        payload: vec![("footnoteRefId".to_owned(), Any::Number(5.0))],
                        attrs: Attrs::new(),
                    },
                    crate::RawOp::InsertEmbed {
                        index: 5,
                        kind: "pilcrow".to_owned(),
                        payload: vec![("paraId".to_owned(), Any::from("body-p"))],
                        attrs: Attrs::new(),
                    },
                ],
                &crate::EditCtx::local("", "2026-07-18T00:00:00Z"),
            )
            .unwrap();
        engine
            .doc()
            .create_story("fn:5", "Footnote text", "Normal", "left")
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {
                "sections": [{
                    "sectionId": "main",
                    "pageSize": {"w": 300, "h": 200},
                    "margins": {"top": 20, "right": 20, "bottom": 20, "left": 20},
                    "noteSettings": {"footnoteColumns": 2}
                }]
            },
            "notes": {"contents": [{"id": 5, "height": 0}]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });

        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let note = &output["layout"]["pages"][0]["noteAreas"][0]["notes"][0];

        assert_eq!(output["notesConverged"], true);
        assert_eq!(
            output["layout"]["pages"][0]["footnoteIds"],
            serde_json::json!([5.0])
        );
        assert_eq!(note["displayLabel"], "1");
        assert_eq!(note["blocks"][0]["runs"][0]["text"], "1  ");
        assert!(note["height"].as_f64().unwrap() > 0.0);
        assert!(
            output["layout"]["pages"][0]["footnoteReservedHeight"]
                .as_f64()
                .unwrap()
                > docx_layout::footnotes::FOOTNOTE_SEPARATOR_HEIGHT
        );
    }

    fn paragraph_pagination_input(first_text: &str, shifted_suffix: bool) -> String {
        let measured: Vec<_> = (0..15)
            .map(|index| {
                let shift = usize::from(shifted_suffix && index > 0);
                let start = index * 2 + shift;
                let text = if index == 0 { first_text } else { "x" };
                serde_json::json!({
                    "block": {
                        "kind": "paragraph",
                        "id": format!("p{index}"),
                        "paraId": format!("para-{index}"),
                        "runs": [{
                            "kind": "text",
                            "text": text,
                            "pmStart": start + 1,
                            "pmEnd": start + 2
                        }],
                        "pmStart": start,
                        "pmEnd": start + 2
                    },
                    "measure": {
                        "kind": "paragraph",
                        "lines": [{
                            "headRun": 0,
                            "headChar": 0,
                            "tailRun": 0,
                            "tailChar": 1,
                            "width": 10,
                            "ascent": 8,
                            "descent": 2,
                            "lineHeight": 20
                        }],
                        "totalHeight": 20
                    }
                })
            })
            .collect();
        serde_json::json!({
            "measured": measured,
            "options": {
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 }
            }
        })
        .to_string()
    }

    #[test]
    fn resident_pagination_reuses_converged_suffix_at_same_positions() {
        let engine = EngineSession::new(14);
        engine
            .layout_document_json(&paragraph_pagination_input("x", false))
            .unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        let caret = engine.resident_caret_snapshot(Some(("p0", 0))).unwrap();
        assert_eq!(caret.frame_epoch, 1);
        assert_eq!(caret.caret_rect.as_ref().unwrap().page_id, "1");
        assert_eq!(caret.caret_rect.as_ref().unwrap().page_index, 0);
        let next_input = paragraph_pagination_input("y", true);
        let incremental = engine.layout_document_json(&next_input).unwrap();
        let full = docx_layout::layout_to_json(&next_input).unwrap();
        assert_eq!(incremental, full);
        engine.build_display_list_frame("{}", 1).unwrap();
        let incremental_display = engine
            .with_display_list(Clone::clone)
            .expect("display list retained");
        let full_display = {
            let pagination = engine.pagination.borrow();
            docx_layout::build_display_list_value_from_resident(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                "{}",
            )
            .unwrap()
        };
        assert_eq!(incremental_display, full_display);

        let stats = engine.stats();
        assert_eq!(stats.pagination_calls, 2);
        assert_eq!(stats.incremental_pagination_calls, 1);
        assert!(stats.pagination_blocks_placed < 30);
        assert!(stats.retained_checkpoints >= 3);
        assert_eq!(stats.rebuilt_pages, 1);
        assert_eq!(stats.incremental_display_builds, 1);
        assert_eq!(stats.rebuilt_display_pages, 4);
    }

    #[test]
    fn resident_dirty_measurement_reuses_host_envelope_and_clean_blocks() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();

        let engine = EngineSession::new(15);
        engine
            .doc()
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        let env = RenderEnv::default();
        let block = engine
            .with_lowered_story("body", &env, |blocks| blocks[0].clone())
            .unwrap();
        let envelope = serde_json::json!({
            "block": block,
            "maxWidth": 180,
            "fontChains": { "liberation sans|0|0": [font_id] },
            "defaults": { "fontSize": 12, "fontFamily": "Liberation Sans" },
            "authoritativeShaping": true
        });
        let extent: ParagraphExtent = serde_json::from_str(
            &engine
                .measure_paragraph_json(&envelope.to_string())
                .unwrap(),
        )
        .unwrap();
        let para_id = block_key(paragraph_identity(&block).unwrap().0).into_owned();
        let initial_input = LayoutInput {
            measured: vec![MeasuredBlock {
                block,
                measure: BlockExtent::Paragraph(extent),
            }],
            options: serde_json::from_value(serde_json::json!({
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 }
            }))
            .unwrap(),
        };
        engine.layout_document_value(initial_input).unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        assert!(engine.can_apply_input("body", &para_id));

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                "!",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let frame = engine.apply_and_layout("body", 1).unwrap();
        assert!(!frame.is_empty());
        assert_eq!(u32::from_le_bytes(frame[12..16].try_into().unwrap()), 0);

        engine
            .doc()
            .delete_range(
                &crate::EditCtx::local("", ""),
                crate::StoryRange::new("body", 5, 6),
            )
            .unwrap();
        let delete_frame = engine.apply_and_layout("body", 2).unwrap();
        assert_eq!(
            u32::from_le_bytes(delete_frame[12..16].try_into().unwrap()),
            0,
            "a resident character deletion must remain a FrameDelta, not full recovery"
        );

        let stats = engine.stats();
        assert_eq!(stats.retained_measure_templates, 1);
        assert_eq!(stats.compatibility_measure_calls, 1);
        assert_eq!(stats.resident_measure_calls, 2);
        assert_eq!(stats.resident_reused_blocks, 0);
        assert_eq!(stats.pagination_calls, 3);
        assert_eq!(stats.incremental_pagination_calls, 2);
        assert_eq!(stats.display_builds, 3);
        docx_layout::clear_measure_fonts();
    }

    /// The profiler clock is host code: re-entering the engine from it must not
    /// abort an edit the document has already committed.
    #[test]
    fn a_reentrant_profiler_clock_does_not_abort_the_edit() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(4242);
        engine
            .doc()
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        let env = RenderEnv::default();
        let block = engine
            .with_lowered_story("body", &env, |blocks| blocks[0].clone())
            .unwrap();
        let extent: ParagraphExtent = serde_json::from_str(
            &engine
                .measure_paragraph_json(
                    &serde_json::json!({
                        "block": block,
                        "maxWidth": 180,
                        "fontChains": { "liberation sans|0|0": [font_id] },
                        "defaults": { "fontSize": 12, "fontFamily": "Liberation Sans" },
                        "authoritativeShaping": true
                    })
                    .to_string(),
                )
                .unwrap(),
        )
        .unwrap();
        engine
            .layout_document_value(LayoutInput {
                measured: vec![MeasuredBlock {
                    block,
                    measure: BlockExtent::Paragraph(extent),
                }],
                options: serde_json::from_value(serde_json::json!({
                    "pageSize": { "w": 200, "h": 120 },
                    "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 }
                }))
                .unwrap(),
            })
            .unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                "!",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();

        let foreign_env = RenderEnv {
            default_tab_stop_twips: Some(720.0),
            ..RenderEnv::default()
        };
        let mut ticks = 0_u32;
        let mut clock = || {
            ticks += 1;
            // Tick 2 is the post-lowering hook; re-lowering under another
            // environment there must not steal the story the edit is reading.
            let env = if ticks == 2 { &foreign_env } else { &env };
            engine.lower_story_json("body", env).unwrap();
            ticks as f64
        };
        let (frame, _) = engine
            .apply_and_layout_profiled("body", 1, &mut clock)
            .unwrap();
        assert!(!frame.is_empty());
        let pagination = engine.pagination.borrow();
        let LayoutBlock::Paragraph(paragraph) =
            &pagination.input.as_ref().unwrap().measured[0].block
        else {
            panic!("paragraph expected");
        };
        assert_eq!(
            paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.default_tab_stop_twips),
            None,
            "the edit keeps the environment it asked for, not the observer's"
        );
        drop(pagination);
        docx_layout::clear_measure_fonts();
    }

    /// The region twin of [`a_reentrant_profiler_clock_does_not_abort_the_edit`].
    #[test]
    fn a_reentrant_profiler_clock_does_not_abort_a_region_edit() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(4243);
        engine
            .doc()
            .create_story("body", "AlphaBravo", "Normal", "left")
            .unwrap();
        engine
            .doc()
            .split_paragraph(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                None,
            )
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        engine.layout_document_with_regions_json(&request).unwrap();
        engine
            .build_display_list_frame(
                &serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0,
            )
            .unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 2),
                "xx",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();

        let env = RenderEnv::default();
        let mut ticks = 0_u32;
        let mut clock = || {
            ticks += 1;
            engine.lower_story_json("body", &env).unwrap();
            ticks as f64
        };
        let (frame, _) = engine
            .apply_and_layout_profiled("body", 1, &mut clock)
            .unwrap();
        assert!(!frame.is_empty());
        docx_layout::clear_measure_fonts();
    }

    /// Rust float equality makes `-0.0 == 0.0`, so a sign-only change to a
    /// layout number keeps its retained measurement. Signed zero measures
    /// identically, so the reuse still matches a full pass.
    #[test]
    fn a_sign_only_zero_change_reuses_a_measurement_matching_the_full_pass() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(777);
        let ctx = crate::EditCtx::local("", "");
        engine
            .doc()
            .create_story(
                "body",
                "Alpha bravo charlie delta echo foxtrot",
                "Normal",
                "left",
            )
            .unwrap();
        engine
            .doc()
            .split_paragraph(&ctx, crate::Position::new("body", 5), None)
            .unwrap();
        let para_ids: Vec<String> = engine
            .with_lowered_story("body", &RenderEnv::default(), |blocks| {
                blocks
                    .iter()
                    .filter_map(|block| {
                        paragraph_identity(block).map(|(id, _)| block_key(id).into_owned())
                    })
                    .collect()
            })
            .unwrap();
        engine
            .doc()
            .set_paragraph_attr(&para_ids[0], "indentLeft", yrs::Any::Number(0.0))
            .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        engine.layout_document_with_regions_json(&request).unwrap();
        engine
            .build_display_list_frame(
                &serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0,
            )
            .unwrap();

        engine
            .doc()
            .set_paragraph_attr(&para_ids[0], "indentLeft", yrs::Any::Number(-0.0))
            .unwrap();
        let lowered = engine
            .with_lowered_story("body", &RenderEnv::default(), |blocks| {
                serde_json::to_string(&blocks[0]).unwrap()
            })
            .unwrap();
        assert!(
            lowered.contains("-0.0"),
            "the sign-only change must reach the lowered block: {lowered}"
        );

        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let after = engine.stats();
        assert_eq!(
            after.resident_measure_calls, before.resident_measure_calls,
            "a sign-only zero change is structurally clean"
        );
        assert_eq!(
            after.resident_reused_blocks,
            before.resident_reused_blocks + 2,
            "both paragraphs reuse their retained extents"
        );
        let fast_json = {
            let pagination = engine.pagination.borrow();
            let regions_state = engine.regions.borrow();
            serialize_region_layout(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                regions_state.as_ref().unwrap().headers_footers.as_ref(),
                true,
            )
            .unwrap()
        };
        let full_json = engine.layout_document_with_regions_json(&request).unwrap();
        assert_eq!(
            fast_json, full_json,
            "the reused measurement matches a full pass"
        );
        docx_layout::clear_measure_fonts();
    }

    /// A paragraph-table-paragraph body for dirty-block detection checks.
    fn resident_walk_fixture() -> String {
        let paragraph = |id: &str, text: &str, start: f64| {
            serde_json::json!({
                "block": {
                    "kind": "paragraph",
                    "id": id,
                    "runs": [{
                        "kind": "text",
                        "text": text,
                        "pmStart": start + 1.0,
                        "pmEnd": start + 2.0
                    }],
                    "pmStart": start,
                    "pmEnd": start + 2.0
                },
                "measure": {
                    "kind": "paragraph",
                    "lines": [{
                        "headRun": 0,
                        "headChar": 0,
                        "tailRun": 0,
                        "tailChar": 1,
                        "width": 10.0,
                        "ascent": 8.0,
                        "descent": 2.0,
                        "lineHeight": 20.0
                    }],
                    "totalHeight": 20.0
                }
            })
        };
        let cell_paragraph = serde_json::json!({
            "kind": "paragraph",
            "id": "c0",
            "runs": [{"kind": "text", "text": "cell", "pmStart": 2.0, "pmEnd": 6.0}],
            "pmStart": 1.0,
            "pmEnd": 6.0
        });
        serde_json::json!({
            "measured": [
                paragraph("p0", "alpha", 0.0),
                {
                    "block": {
                        "kind": "table",
                        "id": "t0",
                        "rows": [{
                            "id": "r0",
                            "cells": [{"id": "c0-cell", "blocks": [cell_paragraph]}]
                        }],
                        "pmStart": 22.0,
                        "pmEnd": 42.0
                    },
                    "measure": {
                        "kind": "table",
                        "rows": [{
                            "cells": [{"blocks": [], "width": 100.0, "height": 20.0}],
                            "height": 20.0
                        }],
                        "columnWidths": [100.0],
                        "totalWidth": 100.0,
                        "totalHeight": 20.0
                    }
                },
                paragraph("p1", "omega", 42.0)
            ],
            "options": {
                "pageSize": {"w": 200.0, "h": 120.0},
                "margins": {"top": 10.0, "right": 10.0, "bottom": 10.0, "left": 10.0}
            }
        })
        .to_string()
    }

    #[test]
    fn resident_walk_marks_only_structurally_changed_blocks_dirty() {
        let engine = EngineSession::new(21);
        engine
            .layout_document_json(&resident_walk_fixture())
            .unwrap();
        let baseline: Vec<LayoutBlock> = {
            let pagination = engine.pagination.borrow();
            pagination
                .input
                .as_ref()
                .unwrap()
                .measured
                .iter()
                .map(|measured| measured.block.clone())
                .collect()
        };
        let walk = |blocks: &[LayoutBlock]| {
            let mut dirty: Vec<(usize, String)> = Vec::new();
            engine
                .resident_layout_input_from_blocks(blocks, &mut |index, key, _, _| {
                    dirty.push((index, key.to_owned()));
                    Ok(BlockExtent::Paragraph(ParagraphExtent {
                        lines: Vec::new(),
                        total_height: 20.0,
                    }))
                })
                .map(|_| dirty)
        };

        assert_eq!(
            walk(&baseline).unwrap(),
            Vec::<(usize, String)>::new(),
            "a no-op relower marks nothing dirty"
        );

        let edited_first = {
            let mut blocks = baseline.clone();
            let LayoutBlock::Paragraph(paragraph) = &mut blocks[0] else {
                panic!("paragraph expected");
            };
            if let Run::Text(text) = &mut paragraph.runs[0] {
                text.text.push('!');
            }
            blocks
        };
        assert_eq!(
            walk(&edited_first).unwrap(),
            vec![(0, "p0".to_owned())],
            "an inserted character dirties exactly its own block"
        );

        let repositioned_tail = {
            let mut blocks = baseline.clone();
            let LayoutBlock::Paragraph(paragraph) = &mut blocks[2] else {
                panic!("paragraph expected");
            };
            paragraph.pm_start = Some(60.0);
            paragraph.pm_end = Some(64.0);
            if let Run::Text(text) = &mut paragraph.runs[0] {
                text.pm_start = Some(61.0);
                text.pm_end = Some(63.0);
            }
            blocks
        };
        assert_eq!(
            walk(&repositioned_tail).unwrap(),
            Vec::<(usize, String)>::new(),
            "fresh absolute positions alone never dirty a block"
        );

        let retitled_cell = {
            let mut blocks = baseline.clone();
            let LayoutBlock::Table(table) = &mut blocks[1] else {
                panic!("table expected");
            };
            let LayoutBlock::Paragraph(cell_paragraph) = &mut table.rows[0].cells[0].blocks[0]
            else {
                panic!("cell paragraph expected");
            };
            if let Run::Text(text) = &mut cell_paragraph.runs[0] {
                text.text.push('!');
            }
            blocks
        };
        let error = walk(&retitled_cell).unwrap_err();
        assert!(error.contains("non-paragraph"), "{error}");

        let repositioned_cell = {
            let mut blocks = baseline.clone();
            let LayoutBlock::Table(table) = &mut blocks[1] else {
                panic!("table expected");
            };
            let LayoutBlock::Paragraph(cell_paragraph) = &mut table.rows[0].cells[0].blocks[0]
            else {
                panic!("cell paragraph expected");
            };
            cell_paragraph.pm_start = Some(30.0);
            cell_paragraph.pm_end = Some(34.0);
            if let Run::Text(text) = &mut cell_paragraph.runs[0] {
                text.pm_start = Some(31.0);
                text.pm_end = Some(35.0);
            }
            blocks
        };
        assert_eq!(
            walk(&repositioned_cell).unwrap(),
            Vec::<(usize, String)>::new(),
            "absolute positions inside a table cell are ignored too"
        );
    }

    #[test]
    fn display_list_is_retained_with_json() {
        let engine = EngineSession::new(17);
        let pagination_input = r#"{
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            }
        }"#;
        let layout: serde_json::Value =
            serde_json::from_str(&engine.layout_document_json(pagination_input).unwrap()).unwrap();
        let display_input = serde_json::json!({
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            },
            "layout": layout,
        })
        .to_string();

        let resident = engine.build_display_list_json(&display_input).unwrap();
        let expected = docx_layout::display_list::build_display_list_json(&display_input).unwrap();
        assert_eq!(resident, expected);
        assert_eq!(engine.stats().frame_epoch, 1);
        assert_eq!(engine.stats().retained_display_pages, 1);
        assert_eq!(engine.stats().retained_display_primitives, 0);
        assert_eq!(engine.stats().display_builds, 1);
        assert_eq!(engine.with_display_list(|list| list.pages.len()), Some(1));
        assert_eq!(
            engine
                .display_hit_test_regions_json(0, 100.0, 100.0)
                .unwrap(),
            r#"{"region":"body","pos":null,"target":"none"}"#
        );
        assert_eq!(engine.display_range_rects_json(0, 1).unwrap(), "[]");
        assert_eq!(
            engine
                .display_range_rects_region_json("body", "", 0, 1)
                .unwrap(),
            "[]"
        );
    }

    #[test]
    fn demo_hit_test_reaches_paragraph_after_table() {
        let engine = EngineSession::new(18);
        crate::seed::seed_from_docx(
            engine.doc(),
            include_bytes!("../../../apps/demo/public/betteroffice-demo.docx"),
        )
        .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
            "measurement": {
                "fontChains": {},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": false
            },
            "renderEnv": {}
        });
        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        engine.build_display_list_json(&output).unwrap();

        let target = engine
            .with_display_list(|list| {
                list.pages
                    .iter()
                    .enumerate()
                    .flat_map(|(page_index, page)| {
                        page.primitives.iter().filter_map(move |primitive| {
                            let docx_layout::display_list::Primitive::Text(text) = primitive else {
                                return None;
                            };
                            (text.text == "Why it matters").then(|| {
                                (
                                    page_index,
                                    text.x.as_f64().unwrap() + text.width.as_f64().unwrap() / 2.0,
                                    text.baseline_y.as_f64().unwrap(),
                                    text.attrs.doc_start.unwrap(),
                                    text.attrs.doc_end.unwrap(),
                                )
                            })
                        })
                    })
                    .next()
            })
            .flatten()
            .expect("demo paragraph after table is positioned");
        let hit: serde_json::Value = serde_json::from_str(
            &engine
                .display_hit_test_regions_json(target.0, target.1, target.2)
                .unwrap(),
        )
        .unwrap();
        let position = hit["pos"].as_i64().expect("body hit has a position");

        assert_eq!(hit["region"], "body");
        assert!((target.3..=target.4).contains(&position));
        assert_eq!(hit["target"], "text");

        // the same line's left margin resolves a position all the same, but is
        // not typeable area
        let margin: serde_json::Value = serde_json::from_str(
            &engine
                .display_hit_test_regions_json(target.0, 8.0, target.2)
                .unwrap(),
        )
        .unwrap();
        assert!(margin["pos"].is_i64());
        assert_eq!(margin["target"], "none");
    }

    /// What a laid-out note area gives a hit probe: the note it carries, and
    /// the two runs every note paints — its presentation label, then its story
    /// text.
    struct PaintedNote {
        kind: Option<String>,
        note_ids: Vec<i64>,
        label_doc_start: Option<i64>,
        text_doc_start: Option<i64>,
        text_doc_end: i64,
        text_center_x: f64,
        text_baseline: f64,
    }

    fn painted_notes(engine: &EngineSession) -> Vec<PaintedNote> {
        engine
            .with_display_list(|list| {
                list.pages[0]
                    .note_areas
                    .iter()
                    .map(|area| {
                        let runs = area
                            .primitives
                            .iter()
                            .filter_map(|primitive| match primitive {
                                docx_layout::display_list::Primitive::Text(run) => Some(run),
                                _ => None,
                            })
                            .collect::<Vec<_>>();
                        let (label, text) = (runs[0], runs[1]);
                        PaintedNote {
                            kind: area.kind.clone(),
                            note_ids: area.note_ids.clone(),
                            label_doc_start: label.attrs.doc_start,
                            text_doc_start: text.attrs.doc_start,
                            text_doc_end: text.attrs.doc_end.expect("note text is positioned"),
                            text_center_x: text.x.as_f64().unwrap()
                                + text.width.as_f64().unwrap() / 2.0,
                            text_baseline: text.baseline_y.as_f64().unwrap(),
                        }
                    })
                    .collect()
            })
            .expect("the display list is built")
    }

    /// Footnotes and endnotes seeded from a real file, laid out and painted
    /// through the resident path, resolve to their own `fn:{id}` / `en:{id}`
    /// stories. The presentation label every note carries has no position of
    /// its own, so this is also where inheriting the body anchor would hand a
    /// click a body position under a note region.
    #[test]
    fn real_notes_resolve_to_their_own_stories() {
        let engine = EngineSession::new(19);
        crate::seed::seed_from_docx(
            engine.doc(),
            include_bytes!("../tests/fixtures/footnote-anchor.docx"),
        )
        .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": {}}]},
            "notes": {"contents": [
                {"id": 2, "noteKind": "footnote", "height": 0},
                {"id": 3, "noteKind": "endnote", "height": 0}
            ]},
            "measurement": {
                "fontChains": {},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": false
            },
            "renderEnv": {}
        });
        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        engine.build_display_list_json(&output).unwrap();

        let painted = painted_notes(&engine);
        assert_eq!(painted.len(), 2, "one area per note kind");

        for (kind, note_id) in [("footnote", 2), ("endnote", 3)] {
            let note = painted
                .iter()
                .find(|note| note.kind.as_deref() == Some(kind))
                .unwrap_or_else(|| panic!("no {kind} area"));
            assert_eq!(note.note_ids, vec![note_id]);
            assert_eq!(
                note.label_doc_start, None,
                "the {kind} label carries a position"
            );
            assert_eq!(
                note.text_doc_start,
                Some(1),
                "the {kind} text lost its story position"
            );

            let hit: serde_json::Value = serde_json::from_str(
                &engine
                    .display_hit_test_regions_json(0, note.text_center_x, note.text_baseline - 2.0)
                    .unwrap(),
            )
            .unwrap();
            assert_eq!(hit["region"], kind);
            assert_eq!(hit["noteId"], note_id);
            assert_eq!(hit["target"], "text");
            let position = hit["pos"].as_i64().expect("the note hit has a position");
            assert!(
                (1..=note.text_doc_end).contains(&position),
                "{kind} hit resolved {position}, outside its story"
            );

            // that story's selection geometry lands in its own area, below the
            // body rects for the very same positions
            let note_rects: serde_json::Value = serde_json::from_str(
                &engine
                    .display_range_rects_region_json(
                        kind,
                        &note_id.to_string(),
                        1,
                        note.text_doc_end,
                    )
                    .unwrap(),
            )
            .unwrap();
            let body_rects: serde_json::Value = serde_json::from_str(
                &engine
                    .display_range_rects_json(1, note.text_doc_end)
                    .unwrap(),
            )
            .unwrap();
            let note_y = note_rects[0]["y"].as_f64().expect("a note rect");
            let body_y = body_rects[0]["y"].as_f64().expect("a body rect");
            assert!(
                note_y > body_y,
                "{kind} rect y {note_y} is not below the body rect y {body_y}"
            );
        }

        // the body reference marks anchoring the notes are still the body's
        let anchor: serde_json::Value = serde_json::from_str(
            &engine
                .display_hit_test_regions_json(0, 100.0, 110.0)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(anchor["region"], "body");
    }

    const LIBERATION: &[u8] =
        include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

    const PACKAGE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

    const DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    fn docx_bytes(styles_body: &str, document_body: &str) -> Vec<u8> {
        let styles = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
{styles_body}
</w:styles>"#
        );
        let document = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>{document_body}</w:body>
</w:document>"#
        );
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
            ("_rels/.rels".to_owned(), PACKAGE_RELS.into()),
            (
                "word/_rels/document.xml.rels".to_owned(),
                DOCUMENT_RELS.into(),
            ),
            ("word/styles.xml".to_owned(), styles.into_bytes()),
            ("word/document.xml".to_owned(), document.into_bytes()),
        ])
        .expect("zip the fixture")
    }

    fn layout_pages(bytes: &[u8], client_id: u64, content_height: f64) -> serde_json::Value {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let engine = EngineSession::new(client_id);
        crate::seed::seed_from_docx(engine.doc(), bytes).unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "options": {
                "pageSize": { "w": 400.0, "h": content_height + 40.0 },
                "margins": { "top": 20.0, "right": 20.0, "bottom": 20.0, "left": 20.0 },
            },
            "regions": { "sections": [{
                "sectionId": "main",
                "pageSize": { "w": 400.0, "h": content_height + 40.0 },
                "margins": { "top": 20.0, "right": 20.0, "bottom": 20.0, "left": 20.0 },
            }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        serde_json::from_str(&output).expect("layout json")
    }

    fn paragraph_slices(layout: &serde_json::Value, block_index: usize) -> Vec<(usize, u64, u64)> {
        let block_id = layout["measured"][block_index]["block"]["id"].clone();
        layout["layout"]["pages"]
            .as_array()
            .expect("pages")
            .iter()
            .enumerate()
            .flat_map(|(page_index, page)| {
                page["fragments"]
                    .as_array()
                    .expect("fragments")
                    .iter()
                    .filter(|fragment| fragment["blockId"] == block_id)
                    .map(move |fragment| {
                        (
                            page_index,
                            fragment["fromLine"].as_u64().unwrap_or(0),
                            fragment["toLine"].as_u64().unwrap_or(0),
                        )
                    })
                    .collect::<Vec<_>>()
            })
            .collect()
    }

    fn line_height(layout: &serde_json::Value, block_index: usize) -> f64 {
        layout["measured"][block_index]["measure"]["lines"][0]["lineHeight"]
            .as_f64()
            .expect("measured line height")
    }

    fn line_count(layout: &serde_json::Value, block_index: usize) -> usize {
        layout["measured"][block_index]["measure"]["lines"]
            .as_array()
            .expect("measured lines")
            .len()
    }

    #[test]
    fn resident_display_falls_back_when_input_adds_a_page() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let engine = EngineSession::new(204);
        let body = format!(
            "{}<w:p><w:r><w:t>Editable paragraph</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(6)
        );
        crate::seed::seed_from_docx(engine.doc(), &docx_bytes("", &body)).unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [{
                "sectionId": "main",
                "properties": {
                    "pageWidth": 4320,
                    "pageHeight": 2880,
                    "marginTop": 300,
                    "marginRight": 300,
                    "marginBottom": 300,
                    "marginLeft": 300
                }
            }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let measured = output["measured"].as_array().unwrap();
        for measured_block in measured {
            let template = serde_json::json!({
                "block": measured_block["block"],
                "maxWidth": 248,
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            });
            engine
                .measure_paragraph_json(&template.to_string())
                .unwrap();
        }
        engine
            .layout_document_json(
                &serde_json::json!({
                    "measured": output["measured"],
                    "options": output["options"]
                })
                .to_string(),
            )
            .unwrap();
        let extras =
            serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
        engine.build_display_list_frame(&extras, 0).unwrap();

        let initial_page_count = engine
            .pagination
            .borrow()
            .layout
            .as_ref()
            .unwrap()
            .pages
            .len();
        let paragraph = engine.doc().paragraphs("body").unwrap().pop().unwrap();
        let insertion = " typed text that makes the editable paragraph wrap";
        let mut offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        let mut frame_epoch = engine.display.borrow().binary_frame_epoch;
        let mut page_count_changed = false;

        for _ in 0..64 {
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", offset),
                    insertion,
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            offset += u32::try_from(insertion.encode_utf16().count()).unwrap();
            let incremental_display_builds = engine.stats().incremental_display_builds;
            engine.apply_and_layout("body", frame_epoch).unwrap();
            frame_epoch = engine.display.borrow().binary_frame_epoch;

            let retained = engine.with_display_list(Clone::clone).unwrap();
            let rebuilt = {
                let pagination = engine.pagination.borrow();
                docx_layout::build_display_list_value_from_resident(
                    pagination.input.as_ref().unwrap(),
                    pagination.layout.as_ref().unwrap(),
                    &extras,
                )
                .unwrap()
            };
            assert_eq!(retained.pages.len(), rebuilt.pages.len());
            for (retained_page, rebuilt_page) in retained.pages.iter().zip(&rebuilt.pages) {
                assert_eq!(retained_page, rebuilt_page);
            }

            if retained.pages.len() > initial_page_count {
                assert!(engine.pagination.borrow().last_incremental);
                assert_eq!(
                    engine.stats().incremental_display_builds,
                    incremental_display_builds
                );
                page_count_changed = true;
                break;
            }
        }

        assert!(page_count_changed, "the edit must add a page");
        docx_layout::clear_measure_fonts();
    }

    const WIDOW_STYLES: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/></w:style>"#;

    fn widow_document(paragraph_ppr: &str) -> String {
        let long = "Widow and orphan control decides where this paragraph may break \
                    across a page boundary, so it has to be long enough to wrap onto \
                    several lines of the narrow page this test lays out.";
        format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Filler</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Body"/>{paragraph_ppr}</w:pPr><w:r><w:t>{long}</w:t></w:r></w:p>"#
        )
    }

    #[test]
    fn authored_widow_control_off_survives_a_docx_round_trip_into_pagination() {
        let inherited = docx_bytes(WIDOW_STYLES, &widow_document(""));
        let probe = layout_pages(&inherited, 201, 4000.0);
        let line = line_height(&probe, 1);
        assert!(line_count(&probe, 1) >= 4, "the rule needs four lines");
        let content_height = line * 2.5;

        let overridden = docx_bytes(WIDOW_STYLES, &widow_document("<w:widowControl/>"));
        let off = layout_pages(&inherited, 202, content_height);
        let on = layout_pages(&overridden, 203, content_height);

        assert_eq!(paragraph_slices(&off, 1)[0], (0, 0, 1));
        assert!(paragraph_slices(&on, 1).iter().all(|(page, ..)| *page > 0));
        assert_eq!(paragraph_slices(&on, 1)[0].1, 0);
    }

    #[test]
    fn negative_top_margin_suppresses_header_expansion_in_resident_regions() {
        use docx_layout::regions::{DocumentRegions, RegionSection};
        let regions = DocumentRegions {
            sections: vec![RegionSection {
                page_size: Some(docx_layout::types::Size {
                    w: 816.0,
                    h: 1056.0,
                }),
                margins: Some(docx_layout::types::PageMargins {
                    top: -1438.0 / 15.0,
                    right: 1797.0 / 15.0,
                    bottom: 96.0,
                    left: 1797.0 / 15.0,
                    header: Some(709.0 / 15.0),
                    footer: Some(48.0),
                }),
                ..Default::default()
            }],
            ..Default::default()
        };
        let mut input: LayoutInput = LayoutInput {
            measured: Vec::new(),
            options: Default::default(),
        };
        let variants = vec![HeaderFooterVariant {
            r_id: "rId1".to_owned(),
            kind: HeaderFooterKind::Header,
            hf_type: HeaderFooterType::Default,
            section_index: 0,
            measured: Vec::new(),
            height: 100.0,
            flow_height: 100.0,
            visual_top: 0.0,
            visual_bottom: 100.0,
            field_widths: Vec::new(),
        }];
        extend_input_for_header_footer(&mut input, &regions, &variants);
        let margins = input.options.margins.expect("extended margins");
        assert_eq!(margins.top, 1438.0 / 15.0);
        assert_eq!(margins.bottom, 96.0);
    }

    #[test]
    fn negative_page_margins_archive_keeps_absolute_geometry_and_signed_save() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let document_body = concat!(
            r#"<w:p><w:r><w:t>Hello negative margins</w:t></w:r></w:p>"#,
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>"#,
            r#"<w:pgMar w:top="-1438" w:right="1797" w:bottom="1440" w:left="1797" w:header="709" w:footer="709"/></w:sectPr>"#
        );
        let bytes = docx_bytes("", document_body);
        let envelope = crate::seed::parse_docx_for_edit(&bytes).unwrap();
        let final_props = envelope
            .document
            .package
            .document
            .final_section_properties
            .clone()
            .expect("final sectPr");
        assert_eq!(final_props.margin_top, Some(-1438.0));
        assert_eq!(final_props.header_distance, Some(709.0));
        let engine = EngineSession::new(911);
        crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
        let properties = serde_json::to_value(&final_props).unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{"sectionId": "main", "properties": properties}]},
            "measurement": {
                "fontChains": {"liberation sans|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let expected_top = 1438.0 / 15.0;
        let fresh: serde_json::Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        assert_eq!(
            fresh["layout"]["pages"][0]["margins"]["top"]
                .as_f64()
                .unwrap(),
            expected_top
        );
        let fresh_y = fresh["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap();
        assert!(fresh_y >= 0.0);
        assert!((fresh_y - expected_top).abs() < 0.05);
        let extras =
            serde_json::json!({"fontChains": {"liberation sans|0|0": [font_id]}}).to_string();
        engine.build_display_list_frame(&extras, 0).unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                " edited",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        engine.apply_and_layout("body", epoch).unwrap();
        let incremental = engine.pagination.borrow();
        let incremental_layout = incremental.layout.as_ref().unwrap();
        assert_eq!(incremental_layout.pages[0].margins.top, expected_top);
        let incremental_y = match &incremental_layout.pages[0].fragments[0] {
            docx_layout::types::Fragment::Paragraph(fragment) => fragment.y,
            other => panic!("expected paragraph fragment, got {other:?}"),
        };
        assert!((incremental_y - expected_top).abs() < 0.05);
        drop(incremental);
        let refreshed: serde_json::Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        let refreshed_top = refreshed["layout"]["pages"][0]["margins"]["top"]
            .as_f64()
            .unwrap();
        let refreshed_y = refreshed["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap();
        assert_eq!(refreshed_top, expected_top);
        assert!((refreshed_y - incremental_y).abs() < 0.001);
        assert!((refreshed_y - expected_top).abs() < 0.05);
        let content = serde_json::to_value(&envelope.document.package.document.content).unwrap();
        let final_json = serde_json::to_value(&final_props).unwrap();
        let relationships =
            serde_json::to_value(&envelope.document.package.relationship_entries).unwrap();
        let save_request: docx_parse::S13SaveRequest = serde_json::from_value(serde_json::json!({
            "determinism": {
                "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
                "now": "2030-01-02T03:04:05.006Z"
            },
            "document": {"content": content, "finalSectionProperties": final_json},
            "relationshipEntries": relationships,
            "options": {"updateModifiedDate": false}
        }))
        .unwrap();
        let saved = docx_parse::write_docx_s13(save_request, &bytes).unwrap();
        let parts = ooxml_opc::unzip_parts(&saved).unwrap();
        let document_xml = String::from_utf8(
            parts
                .iter()
                .find(|(path, _)| path == "word/document.xml")
                .unwrap()
                .1
                .clone(),
        )
        .unwrap();
        assert!(document_xml.contains(r#"w:top="-1438""#));
        let reopened = crate::seed::parse_docx_for_edit(&saved).unwrap();
        assert_eq!(
            reopened
                .document
                .package
                .document
                .final_section_properties
                .as_ref()
                .unwrap()
                .margin_top,
            Some(-1438.0)
        );
        docx_layout::clear_measure_fonts();
    }
}
