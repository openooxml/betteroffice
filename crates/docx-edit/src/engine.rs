//! Resident editor-engine state shared by the wasm facade and native tests.
//!
//! [`EngineSession`] is the migration owner described by
//! `openspec/changes/engine-unification/00-DESIGN.md`: it owns the live
//! [`EditingDoc`] and render-derived state.  The initial migration unit keeps
//! the legacy `LayoutBlock[]` JSON boundary as a parity/debug export, but the
//! lowered Rust values stay resident and are reused for repeated reads of the
//! same document/render generation.

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
use docx_layout::hit::HitRegion;
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
    blocks: Vec<LayoutBlock>,
    /// Legacy `LayoutBlock[]` JSON, built lazily on the first
    /// [`EngineSession::lower_story_json`] read. The resident hot path never
    /// serializes it.
    parity_json: Option<String>,
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

fn region_measurement_widths(
    blocks: &[LayoutBlock],
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
        .iter()
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
        page_height: size.h,
        margin_top: margins.top,
        content_height: size.h - margins.top - margins.bottom,
    }
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
            let margins = docx_layout::section_breaks::resolve_page_margins(
                section.margins.as_ref().or(Some(&fallback_margins)),
            );
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

/// Observability snapshot for parity tests and the opt-in profiler.
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
    let mut value = serde_json::to_value(measured)
        .map_err(|error| format!("fingerprint measured block: {error}"))?;
    strip_absolute_positions(&mut value);
    serde_json::to_vec(&value)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint measured block: {error}"))
}

fn measured_fingerprints(input: &LayoutInput) -> Result<Vec<u64>, String> {
    input.measured.iter().map(measured_fingerprint).collect()
}

fn block_fingerprint(block: &LayoutBlock) -> Result<u64, String> {
    let mut value = serde_json::to_value(block)
        .map_err(|error| format!("fingerprint layout block: {error}"))?;
    strip_absolute_positions(&mut value);
    serde_json::to_vec(&value)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint layout block: {error}"))
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

fn block_key(id: &BlockId) -> String {
    match id {
        BlockId::Num(value) if value.fract() == 0.0 => format!("{}", *value as i64),
        BlockId::Num(value) => value.to_string(),
        BlockId::Str(value) => value.clone(),
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

fn position_deltas(previous: &LayoutInput, next: &LayoutInput) -> HashMap<String, i64> {
    previous
        .measured
        .iter()
        .zip(&next.measured)
        .filter_map(|(previous, next)| {
            let (previous_id, previous_start) = paragraph_identity(&previous.block)?;
            let (next_id, next_start) = paragraph_identity(&next.block)?;
            let key = block_key(previous_id);
            if key != block_key(next_id) {
                return None;
            }
            let delta = next_start? as i64 - previous_start? as i64;
            (delta != 0).then_some((key, delta))
        })
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
            .all(|(previous, next)| {
                paragraph_identity(&previous.block)
                    .zip(paragraph_identity(&next.block))
                    .is_some_and(|((previous_id, _), (next_id, _))| {
                        block_key(previous_id) == block_key(next_id)
                    })
            })
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

    /// Internal editing surface used by the compatibility wasm facade.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn doc(&self) -> &EditingDoc {
        &self.doc
    }

    pub fn doc_epoch(&self) -> u64 {
        self.doc_epoch.get()
    }

    /// Resident lowering path. The returned slice is retained behind the
    /// engine session; callers that still need ownership use
    /// [`Self::lower_story_json`] during migration.
    pub fn with_lowered_story<T>(
        &self,
        story: &str,
        env: &RenderEnv,
        read: impl FnOnce(&[LayoutBlock]) -> T,
    ) -> Result<T, BridgeError> {
        let epoch = self.doc_epoch();
        let is_hit = self
            .render
            .borrow()
            .stories
            .get(story)
            .is_some_and(|cached| cached.doc_epoch == epoch && cached.env == *env);

        if !is_hit {
            let blocks = yrs_doc_to_layout_blocks(&self.doc, story, env)?;
            let mut render = self.render.borrow_mut();
            render.cache_misses = render.cache_misses.wrapping_add(1);
            render.stories.insert(
                story.to_owned(),
                LoweredStory {
                    doc_epoch: epoch,
                    env: env.clone(),
                    blocks,
                    parity_json: None,
                },
            );
        } else {
            let mut render = self.render.borrow_mut();
            render.cache_hits = render.cache_hits.wrapping_add(1);
        }

        let render = self.render.borrow();
        let cached = render
            .stories
            .get(story)
            .expect("resident story exists after lowering");
        Ok(read(&cached.blocks))
    }

    /// Compatibility/parity export for the pre-engine host. Production can
    /// stop consuming this once layout and display state move into the
    /// session; keeping it byte-identical makes the migration gate explicit.
    pub fn lower_story_json(&self, story: &str, env: &RenderEnv) -> Result<String, BridgeError> {
        self.with_lowered_story(story, env, |_| ())?;
        let mut render = self.render.borrow_mut();
        let lowered = render
            .stories
            .get_mut(story)
            .expect("resident story exists after lowering");
        Ok(lowered
            .parity_json
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
        let previous_fingerprint = block_fingerprint(previous_block).ok()?;
        measurement.templates.values().find_map(|template| {
            if !template.resident_safe {
                return None;
            }
            let block: LayoutBlock =
                serde_json::from_value(template.envelope.get("block")?.clone()).ok()?;
            (block_fingerprint(&block).ok()? == previous_fingerprint)
                .then(|| template.envelope.clone())
        })
    }

    /// Parse, paginate, and retain the measured/options input and resulting
    /// Layout. The JSON result is the migration oracle until the binary frame
    /// cutover consumes the retained typed values directly.
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
            let render_env: RenderEnv = serde_json::from_value(render_env)
                .map_err(|error| format!("parse render environment: {error}"))?;
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

    /// The full region pass minus the JSON envelope: pagination, note
    /// stabilization, and header/footer measurement all land in retained
    /// state. `apply_input`'s fallback consumes this directly so a keystroke
    /// never serializes a layout nobody reads. Returns `notes_converged`.
    fn layout_document_with_regions_value(&self, input_json: &str) -> Result<bool, String> {
        let request: RegionLayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        let (mut input, regions, mut notes, measurement, render_env, body_story) = request.split();
        let parsed_render_env = if render_env.is_null() {
            None
        } else {
            Some(
                serde_json::from_value::<RenderEnv>(render_env)
                    .map_err(|error| format!("parse render environment: {error}"))?,
            )
        };
        let resident_body = body_story.is_some();
        if let Some(story) = body_story.as_deref() {
            let render_env = parsed_render_env
                .as_ref()
                .ok_or_else(|| "resident body layout requires a render environment".to_owned())?;
            let mut blocks = self
                .with_lowered_story(story, render_env, <[LayoutBlock]>::to_vec)
                .map_err(|error| error.to_string())?;
            apply_section_geometry_to_blocks(&mut blocks, &mut input.options, &regions);
            let widths = region_measurement_widths(&blocks, &input, &regions);
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
        } else {
            apply_section_geometry(&mut input, &regions);
        }
        let mut measured_headers_footers = if let Some(render_env) = parsed_render_env.as_ref() {
            self.measure_header_footer_payload(&mut input, &regions, &measurement, render_env)?
        } else {
            None
        };
        self.layout_document_value(input.clone())?;
        let mut initial_layout = self
            .pagination
            .borrow()
            .layout
            .clone()
            .expect("layout retained after successful pagination");
        apply_document_regions(&mut initial_layout, &regions);
        let refs = input
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
        let base_input = input.clone();
        let stabilized = stabilize_note_layout(
            |reserved| {
                let mut pass = base_input.clone();
                pass.options.footnote_reserved_heights = reservation_options(reserved);
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
            input.options.footnote_reserved_heights =
                reservation_options(&stabilized.reserved_heights);
            self.layout_document_value(input)?;
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
                serde_json::to_value(payload)
                    .map_err(|error| format!("serialize headers/footers: {error}"))
            })
            .transpose()?;
        let headers_footers = measured_headers_footers
            .or_else(|| regions.headers_footers.clone());
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
                notes_clear,
            }),
        }));
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
                let blocks = self
                    .with_lowered_story(&format!("hf:{r_id}"), render_env, <[LayoutBlock]>::to_vec)
                    .map_err(|error| error.to_string())?;
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
            let previous = self.pagination.borrow();
            let first_dirty = previous
                .block_fingerprints
                .iter()
                .zip(&block_fingerprints)
                .position(|(previous, next)| previous != next);
            if let Some(dirty_index) = first_dirty
                && incremental_eligible(&previous, &input, input_options_fingerprint)
            {
                let previous_input = previous.input.as_ref().expect("eligibility checked input");
                deltas = position_deltas(previous_input, &input);
                let attempted = docx_layout::place::layout_document_incremental(
                    &mut input,
                    previous
                        .layout
                        .as_ref()
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
        let blocks = self
            .with_lowered_story(story, &env, <[LayoutBlock]>::to_vec)
            .map_err(|error| error.to_string())?;
        after_lower();
        self.resident_layout_input_from_blocks(blocks, &mut |_, key, previous_block, next_block| {
            let mut envelope = self
                .measurement_envelope_for_block(key, previous_block)
                .ok_or_else(|| {
                    format!("resident measurement template missing for block {key:?}")
                })?;
            let fields = envelope
                .as_object_mut()
                .ok_or_else(|| "resident measurement envelope is not an object".to_owned())?;
            fields.insert(
                "block".to_owned(),
                serde_json::to_value(&*next_block)
                    .map_err(|error| format!("serialize dirty paragraph: {error}"))?,
            );
            let envelope_json = serde_json::to_string(&envelope)
                .map_err(|error| format!("serialize measurement envelope: {error}"))?;
            let extent_json = docx_layout::measure_paragraph_json_resident(&envelope_json)?;
            let extent: ParagraphExtent = serde_json::from_str(&extent_json)
                .map_err(|error| format!("parse resident paragraph extent: {error}"))?;
            Ok(BlockExtent::Paragraph(extent))
        })
    }

    /// Shared dirty-block walk over a freshly lowered story: fingerprint-clean
    /// blocks reuse their retained extents (with fresh absolute positions),
    /// changed paragraph blocks are re-measured through `measure_dirty`
    /// (`(block_index, block_key, previous_block, next_block) -> extent`).
    fn resident_layout_input_from_blocks(
        &self,
        blocks: Vec<LayoutBlock>,
        measure_dirty: &mut dyn FnMut(
            usize,
            &str,
            &LayoutBlock,
            &mut LayoutBlock,
        ) -> Result<BlockExtent, String>,
    ) -> Result<ResidentLayoutInput, String> {
        let (previous, previous_fingerprints) = {
            let pagination = self.pagination.borrow();
            (
                pagination
                    .input
                    .clone()
                    .ok_or_else(|| "resident pagination input is not built".to_owned())?,
                pagination.block_fingerprints.clone(),
            )
        };
        let paragraph_merge = blocks.len().checked_add(1) == Some(previous.measured.len());
        if blocks.len() != previous.measured.len() && !paragraph_merge {
            return Err("resident plain-text input changed the block structure".to_owned());
        }
        if previous.measured.len() != previous_fingerprints.len() {
            return Err("resident pagination fingerprints are not built".to_owned());
        }

        let mut previous_blocks = previous
            .measured
            .into_iter()
            .zip(previous_fingerprints)
            .peekable();
        let mut skipped_merged_paragraph = false;
        let mut measured = Vec::with_capacity(blocks.len());
        let mut block_fingerprints = Vec::with_capacity(blocks.len());
        let mut resident_measure_calls = 0_u64;
        let mut resident_reused_blocks = 0_u64;
        for (block_index, mut next_block) in blocks.into_iter().enumerate() {
            let mut previous_entry = previous_blocks.next().ok_or_else(|| {
                "resident plain-text input changed the block structure".to_owned()
            })?;
            if paragraph_merge && !resident_block_slots_match(&previous_entry.0.block, &next_block)
            {
                if skipped_merged_paragraph || paragraph_identity(&previous_entry.0.block).is_none()
                {
                    return Err("resident plain-text input changed the block structure".to_owned());
                }
                skipped_merged_paragraph = true;
                previous_entry = previous_blocks.next().ok_or_else(|| {
                    "resident plain-text input changed the block structure".to_owned()
                })?;
            }
            if paragraph_merge && !resident_block_slots_match(&previous_entry.0.block, &next_block)
            {
                return Err("resident plain-text input changed stable block identity".to_owned());
            }
            let (previous_measured, previous_fingerprint) = previous_entry;
            let (Some((next_id, _)), Some((previous_id, _))) = (
                paragraph_identity(&next_block),
                paragraph_identity(&previous_measured.block),
            ) else {
                if block_fingerprint(&next_block)? != block_fingerprint(&previous_measured.block)? {
                    return Err(
                        "resident plain-text input changed a non-paragraph block".to_owned()
                    );
                }
                measured.push(MeasuredBlock {
                    block: next_block,
                    measure: previous_measured.measure,
                });
                block_fingerprints.push(previous_fingerprint);
                resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
                continue;
            };
            let key = block_key(next_id);
            if key != block_key(previous_id) {
                return Err("resident plain-text input changed stable block identity".to_owned());
            }
            if block_fingerprint(&next_block)? == block_fingerprint(&previous_measured.block)? {
                measured.push(MeasuredBlock {
                    block: next_block,
                    measure: previous_measured.measure,
                });
                block_fingerprints.push(previous_fingerprint);
                resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
                continue;
            }

            let measure =
                measure_dirty(block_index, &key, &previous_measured.block, &mut next_block)?;
            let measured_block = MeasuredBlock {
                block: next_block,
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
                options: previous.options,
            },
            block_fingerprints,
        })
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
                fast.notes_clear
                    .then(|| (Rc::clone(&fast.regions), Rc::clone(&fast.measurement)))
            })
        };
        let Some((regions, measurement)) = fast_config else {
            return Ok(false);
        };
        let env = {
            let render = self.render.borrow();
            let Some(lowered) = render.stories.get(story) else {
                return Ok(false);
            };
            lowered.env.clone()
        };
        let blocks = self
            .with_lowered_story(story, &env, <[LayoutBlock]>::to_vec)
            .map_err(|error| error.to_string())?;
        phase(RegionResidentPhase::Lowered);
        let (widths, geometry, previous_pages) = {
            let pagination = self.pagination.borrow();
            let (Some(input), Some(layout)) =
                (pagination.input.as_ref(), pagination.layout.as_ref())
            else {
                return Ok(false);
            };
            (
                region_measurement_widths(&blocks, input, &regions),
                initial_float_page_geometry(input, &regions),
                layout.pages.len(),
            )
        };
        let default_width = widths.first().copied().unwrap_or(0.0);
        if docx_layout::measure_blocks::has_floating_zones(
            &blocks,
            default_width,
            measurement.as_ref(),
            Some(&geometry),
        )? {
            return Ok(false);
        }
        let resident = match self.resident_layout_input_from_blocks(
            blocks,
            &mut |index, _key, _previous_block, next_block| {
                let width = widths.get(index).copied().unwrap_or(default_width);
                docx_layout::measure_blocks::measure_block(next_block, width, measurement.as_ref())
            },
        ) {
            Ok(resident) => resident,
            // Structural mismatch (beyond the supported plain-edit shapes) —
            // the full pass handles it.
            Err(_) => return Ok(false),
        };
        phase(RegionResidentPhase::Measured);
        self.layout_document_value_with_fingerprints(resident.input, resident.block_fingerprints)?;
        let mut pagination = self.pagination.borrow_mut();
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

    /// Build and retain the typed display list. The JSON result remains the
    /// byte-for-byte production bridge until FrameDelta becomes the only
    /// browser render input.
    pub fn build_display_list_json(&self, input_json: &str) -> Result<String, String> {
        let list = docx_layout::build_display_list_value(input_json)?;
        let parity_json =
            serde_json::to_string(&list).map_err(|error| format!("serialize: {error}"))?;
        let mut display = self.display.borrow_mut();
        display.list = Some(list);
        display.resident_input = None;
        display.frame_epoch = display.frame_epoch.wrapping_add(1);
        display.display_builds = display.display_builds.wrapping_add(1);
        Ok(parity_json)
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
        let previous_pages = display.pages.clone();
        let mut next_page_id = display.next_page_id;
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
                    &previous_pages,
                    epochs,
                    &mut next_page_id,
                    rebuilt_page_start..rebuilt_page_end,
                )?
            } else {
                encode_frame_delta(list, &previous_pages, epochs, full, &mut next_page_id)?
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

    /// Body range geometry directly against the resident display list.
    pub fn display_range_rects_json(&self, from: i64, to: i64) -> Result<String, String> {
        self.with_display_list(|list| docx_layout::hit::range_rects(list, from, to))
            .ok_or_else(|| "resident display list is not built".to_owned())
            .and_then(|rects| serde_json::to_string(&rects).map_err(|error| error.to_string()))
    }

    /// Header/footer/body range geometry directly against the resident list.
    pub fn display_range_rects_region_json(
        &self,
        region: &str,
        r_id: &str,
        from: i64,
        to: i64,
    ) -> Result<String, String> {
        let region = match region {
            "body" => HitRegion::Body,
            "header" => HitRegion::Header,
            "footer" => HitRegion::Footer,
            other => return Err(format!("unknown region {other:?}")),
        };
        let r_id = (!r_id.is_empty()).then_some(r_id);
        self.with_display_list(|list| {
            docx_layout::hit::range_rects_in_region(list, region, r_id, from, to)
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
        let oracle = serde_json::to_string(
            &crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap(),
        )
        .unwrap();
        assert_eq!(
            first, oracle,
            "resident output must match the legacy oracle"
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
    fn pagination_input_and_layout_are_retained_with_parity_json() {
        let engine = EngineSession::new(13);
        let input = r#"{
            "measured": [],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}
            }
        }"#;
        let resident = engine.layout_document_json(input).unwrap();
        let oracle = docx_layout::layout_to_json(input).unwrap();
        assert_eq!(resident, oracle);
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
    fn region_layout_operation_stabilizes_and_emits_note_areas() {
        let engine = EngineSession::new(132);
        let request = serde_json::json!({
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
                "contents": [{"id": 7, "height": 20}]
            }
        });

        let output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
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

    /// Perf probe, not a correctness test. Compares the resident region fast
    /// path against the pre-existing full-pass-per-keystroke behavior on a
    /// paragraph-heavy document.
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
        engine.layout_document_with_regions_json(&request).unwrap();
        engine.build_display_list_frame(&extras, 0).unwrap();

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
            // the pre-fast-path region branch: full pass + serialized envelope
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
    fn resident_pagination_reuses_converged_suffix_with_position_parity() {
        let engine = EngineSession::new(14);
        engine
            .layout_document_json(&paragraph_pagination_input("x", false))
            .unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
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
        let para_id = block_key(paragraph_identity(&block).unwrap().0);
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

    #[test]
    fn display_list_is_retained_with_parity_json() {
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
        let oracle = docx_layout::display_list::build_display_list_json(&display_input).unwrap();
        assert_eq!(resident, oracle);
        assert_eq!(engine.stats().frame_epoch, 1);
        assert_eq!(engine.stats().retained_display_pages, 1);
        assert_eq!(engine.stats().retained_display_primitives, 0);
        assert_eq!(engine.stats().display_builds, 1);
        assert_eq!(engine.with_display_list(|list| list.pages.len()), Some(1));
        assert_eq!(
            engine
                .display_hit_test_regions_json(0, 100.0, 100.0)
                .unwrap(),
            r#"{"region":"body","pos":null}"#
        );
        assert_eq!(engine.display_range_rects_json(0, 1).unwrap(), "[]");
        assert_eq!(
            engine
                .display_range_rects_region_json("body", "", 0, 1)
                .unwrap(),
            "[]"
        );
    }
}
