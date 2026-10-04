//! Resident editor-engine state.

use std::borrow::Cow;
use std::cell::{Cell, RefCell};
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::hash::{DefaultHasher, Hasher as _};
use std::rc::Rc;
use std::sync::Arc;

use docx_layout::display_list::DisplayList;
use docx_layout::footnotes::{
    FOOTNOTE_COLUMN_GAP_PX, NoteKind, NoteSeparatorHeights, OrderedMap, apply_note_presentation,
    assign_note_presentations, attach_note_areas, build_note_presentations,
    calculate_note_reserved_heights, collect_note_refs, map_note_anchors_to_pages,
    map_notes_to_pages, reservation_surplus_pages, stabilize_note_layout, stamp_note_pages,
};
use docx_layout::header_footer::{
    HeaderFooterKind, HeaderFooterMetrics, HeaderFooterPayload, HeaderFooterType,
    HeaderFooterVariant, extend_body_margins, header_footer_float_bands, measure_header_footer,
    resolve_header_footer_field_widths,
};
use docx_layout::hit::{CaretRect, VerticalDirection};
use docx_layout::paragraph_spacing::resolve_doc_grid_pitch;
use docx_layout::paragraph_spacing::resolve_line_unit_spacing;
use docx_layout::place::LayoutCheckpoint;
use docx_layout::regions::{
    DocumentRegions, RegionLayoutInput, apply_document_regions_tracked, apply_section_geometry,
    apply_section_geometry_to_blocks, effective_header_footer_refs,
};
use docx_layout::types::{
    BlockExtent, BlockId, ColumnLayout, Fragment, Input as LayoutInput, Layout, LayoutBlock,
    MeasuredBlock, NoteAreaContract, ParagraphExtent, Run, SectionBreakType, SectionPageMargins,
};
use ooxml_text::measure::{FontChainDependencies, FontChains};
use serde::Serialize;
use yrs::{StickyIndex, Subscription, Transact};

use crate::EditingDoc;
use crate::bridge::{BridgeError, LoweringMap, RenderEnv, RevisionPreview};
use crate::fingerprint::Fingerprint;
use crate::frame_delta::{
    DisplayChanges, FrameEpochs, FramePageSnapshot, PageShiftRun, encode_frame_delta,
    encode_frame_delta_changes,
};
use crate::structured::pages::{self, PageLimits};
use crate::structured::{
    AnchorScope, DocxLayoutMap, DocxPagedStructuredContent, DocxSnapshotLayoutMap, ExportFailure,
    ExportFailureCode, ExportRead, ExportRefusal, PageExportOptions, RevisionView, StoryKind,
};

#[derive(Debug)]
struct LoweredStory {
    doc_epoch: u64,
    env: RenderEnv,
    /// The document's media sources the lowering read.
    media: crate::media::MediaSources,
    /// Shared so a reader can hold the lowering it asked for without the cache
    /// borrow, and without copying the story.
    blocks: Rc<Vec<LayoutBlock>>,
    /// Where the blocks' positions came from, recorded by the same lowering.
    map: Rc<LoweringMap>,
    /// Blocks the lowering left out that a revision preview can reveal.
    revealable_blocks: Rc<Vec<LayoutBlock>>,
    /// Lazily serialized layout blocks.
    serialized_blocks: Option<String>,
    local: crate::bridge::local::LocalLowering,
    preview: Option<Rc<crate::bridge::preview::PreviewUnits>>,
}

#[derive(Debug, Default)]
struct RenderState {
    stories: HashMap<String, LoweredStory>,
    cache_hits: u64,
    cache_misses: u64,
    preview_patches: u64,
    preview_fallbacks: u64,
}

struct LoweredNoteSeparators {
    state: Arc<[u8]>,
    doc: EditingDoc,
    env: RenderEnv,
    blocks: HashMap<String, Rc<Vec<LayoutBlock>>>,
    revealable: HashMap<String, Rc<Vec<LayoutBlock>>>,
}

#[derive(Debug)]
struct PreviewFontRequirements {
    doc_epoch: u64,
    request_fingerprint: u64,
    /// `None` when the superset needs script fallbacks and each preview takes the exact path.
    json: Option<String>,
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
    /// [`layout_options_fingerprint`] of `request_json`.
    request_fingerprint: String,
    headers_footers: Option<serde_json::Value>,
    notes_converged: bool,
    provisional: bool,
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
    cached_page_totals: bool,
    regions: Rc<DocumentRegions>,
    measurement: Rc<docx_layout::measure_blocks::MeasurementConfig>,
    /// Non-font measurement config and font-store identity hashed once.
    measurement_fingerprint: u64,
    /// The measurement fonts' generation the fast path may keep measuring with.
    fonts: (u64, usize),
    /// The header and footer stories the retained bands were measured from, fingerprinted
    /// by [`EngineSession::regional_fingerprint`]; the fast path keeps those bands only
    /// while it is unchanged.
    regional: u64,
    /// True when the last full pass had no normal footnote/endnote contents
    /// and no note references — the fast path skips note stabilization
    /// entirely, so it requires a note-free document.
    notes_clear: bool,
    /// The environment the pass lowered the body with. A body lowered since
    /// with another one, as by a region layout begun and then abandoned, is not
    /// the pass's.
    render_env: RenderEnv,
}

/// What a completed region layout of the session's own stories was computed from, published
/// together with it so a paged export can tell whether the retained layout still describes the
/// document. Only a pass that lowered the body, headers, footers and notes from the session
/// publishes one.
#[derive(Debug)]
struct LayoutCapture {
    /// The document version the layout lowered.
    version: crate::batch::DocumentVersion,
    /// The pagination serial of the layout.
    serial: u64,
    /// The generation of the measurement fonts the layout measured with.
    fonts: (u64, usize),
    note_settlement: NoteSettlement,
    /// The environment every story of the pass was lowered with.
    render_env: RenderEnv,
    headers_footers: Option<Rc<HeaderFooterPayload>>,
    notes: Rc<Vec<docx_layout::footnotes::NoteContent>>,
}

/// How the note reservations of a layout pass settled.
#[derive(Debug)]
enum NoteSettlement {
    /// The reservations are a fixed point of the layout.
    Converged,
    /// The note passes alternated, and the layout keeps a reservation that covers every page's
    /// notes; these page indexes reserve more than their notes take.
    Covering(Vec<usize>),
    /// Some page's notes take more space than the layout reserves for them.
    Unsettled,
}

struct RegionPass {
    notes_converged: bool,
    /// The layout covers only a leading part of the body.
    provisional: bool,
}

/// A region layout pass up to the end of body measurement, which
/// [`EngineSession::finish_region_layout`] completes.
struct PreparedRegionLayout {
    input_json: String,
    request_fingerprint: String,
    input: LayoutInput,
    regions: DocumentRegions,
    notes: docx_layout::footnotes::NoteLayoutInput,
    measurement: docx_layout::measure_blocks::MeasurementConfig,
    parsed_render_env: Option<RenderEnv>,
    revision_preview_key: u64,
    measurement_fingerprint: u64,
    fonts: (u64, usize),
    resident_body: bool,
    /// Whether the pass lowers the `body` story, the one preview changes are traced in.
    main_body: bool,
    block_fingerprints: Option<Vec<Fingerprint>>,
    lowered_from: Option<Rc<Vec<LayoutBlock>>>,
    input_lowering: Option<(u64, Rc<LoweringMap>)>,
    has_floats: bool,
    measured_widths: Vec<f64>,
    measured_font_dependencies: Vec<FontChainDependencies>,
    measured_table_wrap_frames: Vec<bool>,
    measured_float_geometry: Option<[f64; 5]>,
    provisional: bool,
    cached_page_totals: bool,
    /// Body blocks still being measured; `input.measured` is final without them.
    body: Option<BodyMeasure>,
}

struct BodyMeasure {
    blocks: Vec<LayoutBlock>,
    widths: Vec<f64>,
    flow: docx_layout::measure_blocks::FloatFlow,
    /// Fingerprints of the measured blocks, taken as they are measured.
    fingerprints: Vec<Fingerprint>,
}

impl PreparedRegionLayout {
    /// Measures up to `blocks` more body blocks; true once the body is measured.
    fn measure(&mut self, blocks: usize) -> Result<bool, String> {
        let Some(body) = self.body.as_mut() else {
            return Ok(true);
        };
        let end = body.flow.measured().saturating_add(blocks);
        body.flow
            .measure_until(&mut body.blocks, &body.widths, &self.measurement, end)?;
        for (block, measure) in body.blocks[body.fingerprints.len()..]
            .iter()
            .zip(&body.flow.extents()[body.fingerprints.len()..])
        {
            body.fingerprints
                .push(measured_parts_fingerprint(block, measure)?);
        }
        if body.flow.measured() < body.blocks.len() {
            return Ok(false);
        }
        let body = self.body.take().expect("body measure present");
        // Header and footer measurement later widens only the section breaks,
        // which the pass fingerprints again.
        self.block_fingerprints = Some(body.fingerprints);
        self.measured_font_dependencies = body.flow.font_dependencies().to_vec();
        self.input.measured = body
            .blocks
            .into_iter()
            .zip(body.flow.into_extents())
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect();
        Ok(true)
    }
}

/// A region layout left between two measurement steps, valid only while the
/// document and the measurement fonts are as they were when it began.
struct ResumableRegionLayout {
    version: crate::batch::DocumentVersion,
    prepared: PreparedRegionLayout,
}

/// How far a resumable region layout has come.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RegionLayoutProgress {
    pub measured_blocks: usize,
    pub body_blocks: usize,
    /// The retained layout JSON, once the pass is complete.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_json: Option<String>,
}

#[derive(Debug)]
struct ResidentLayoutInput {
    input: LayoutInput,
    block_fingerprints: Vec<Fingerprint>,
    font_dependencies: Vec<FontChainDependencies>,
    /// The doc epoch and lowering map `input` was built from.
    lowering: Option<(u64, Rc<LoweringMap>)>,
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

fn region_measurement_frames<'a>(
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
    input: &LayoutInput,
    regions: &DocumentRegions,
) -> (Vec<f64>, Vec<bool>) {
    let blocks: Vec<_> = blocks.into_iter().collect();
    let section_breaks: Vec<_> = blocks
        .iter()
        .copied()
        .filter_map(|block| match block {
            LayoutBlock::SectionBreak(section_break) => Some(section_break),
            _ => None,
        })
        .collect();
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
    let mut previous_section = None;
    let mut placement_page_width = fallback_size.w;
    let mut placement_margins = &fallback_margins;
    let mut deferred_width = false;
    let negative_flows = docx_layout::measure_blocks::negative_indent_float_flows(&blocks);
    blocks
        .into_iter()
        .zip(negative_flows)
        .map(|(block, negative_flow)| {
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
            let section_break = section_breaks.get(section_index);
            let placement_columns = match section_break {
                Some(section_break) => section
                    .and_then(|section| section.columns.as_ref())
                    .or(section_break.columns.as_ref()),
                None => input.options.columns.as_ref(),
            };
            if previous_section != Some(section_index) {
                let previous_content_width =
                    placement_page_width - placement_margins.left - placement_margins.right;
                let (placement_size, margins) = match section_break {
                    Some(section_break) => (
                        section
                            .and_then(|section| section.page_size.as_ref())
                            .or(section_break.page_size.as_ref()),
                        section
                            .and_then(|section| section.margins.as_ref())
                            .or(section_break.margins.as_ref()),
                    ),
                    None => (
                        Some(
                            input
                                .options
                                .final_page_size
                                .as_ref()
                                .unwrap_or(&fallback_size),
                        ),
                        Some(
                            input
                                .options
                                .final_margins
                                .as_ref()
                                .unwrap_or(&fallback_margins),
                        ),
                    ),
                };
                if let Some(size) = placement_size {
                    placement_page_width = size.w;
                }
                if let Some(margins) = margins {
                    placement_margins = margins;
                }
                let content_width =
                    placement_page_width - placement_margins.left - placement_margins.right;
                let break_type = match section_break {
                    Some(section_break) => section_break.break_type,
                    None => input.options.body_break_type,
                }
                .or_else(|| {
                    section_index
                        .checked_sub(1)
                        .and_then(|index| section_breaks.get(index))
                        .and_then(|section_break| section_break.break_type)
                });
                let continuous = break_type == Some(SectionBreakType::Continuous)
                    || section.and_then(|section| section.section_start)
                        == Some(SectionBreakType::Continuous);
                deferred_width = section_index > 0
                    && continuous
                    && (deferred_width || previous_content_width != content_width);
                previous_section = Some(section_index);
            }
            let table_wrap_frame = regions.sections.get(section_index).is_some()
                && columns.is_none_or(|columns| columns.count == 1.0)
                && placement_columns.is_none_or(|columns| columns.count == 1.0)
                && width == placement_page_width - placement_margins.left - placement_margins.right
                && !deferred_width
                && !negative_flow;
            if matches!(block, LayoutBlock::SectionBreak(_)) {
                section_index += 1;
            }
            (width, table_wrap_frame)
        })
        .unzip()
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

/// A block index a measured prefix may end at. The extent of a bare paragraph
/// mark depends on whether a section break follows it, so no prefix ends just
/// before one, or just before a mark that precedes one. A keep-with-next run
/// is placed by its height through its follower, so no prefix ends inside one.
fn prefix_boundary(blocks: &[LayoutBlock], mut end: usize) -> usize {
    let breaks_at = |index: usize| matches!(blocks.get(index), Some(LayoutBlock::SectionBreak(_)));
    let keeps_with_next = |index: usize| {
        matches!(
            blocks.get(index),
            Some(LayoutBlock::Paragraph(paragraph))
                if paragraph.attrs.as_ref().and_then(|attrs| attrs.keep_next) == Some(true)
        )
    };
    while end < blocks.len()
        && (breaks_at(end) || breaks_at(end + 1) || (end > 0 && keeps_with_next(end - 1)))
    {
        end += 1;
    }
    end
}

fn section_breaks(blocks: &[LayoutBlock]) -> usize {
    blocks
        .iter()
        .filter(|block| matches!(block, LayoutBlock::SectionBreak(_)))
        .count()
}

/// The layout options of the document cut after section `last`, which then
/// lays out as the final section.
fn options_through_section(
    request: &docx_layout::types::LayoutOptions,
    regions: &DocumentRegions,
    last: usize,
) -> docx_layout::types::LayoutOptions {
    let through = DocumentRegions {
        sections: regions.sections.iter().take(last + 1).cloned().collect(),
        even_and_odd_headers: regions.even_and_odd_headers,
        ..DocumentRegions::default()
    };
    let mut options = request.clone();
    apply_section_geometry_to_blocks::<LayoutBlock>(&mut [], &mut options, &through);
    options
}

/// Measures leading blocks until their pagination runs two pages past
/// `pages` outside a section with columns, far enough that no later block
/// moves the first `pages` pages.
/// Returns the extents of the measured prefix; all blocks when it never does.
/// With `anchored` objects the prefix is measured whole each time it grows: a
/// float applies from its anchor on, and zones are extracted at the body's
/// first width, which a batch opening in a later section would not use.
#[allow(clippy::too_many_arguments)]
fn measure_page_prefix(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    table_wrap_frames: &[bool],
    measurement: &docx_layout::measure_blocks::MeasurementConfig,
    geometry: &docx_layout::measure_blocks::FloatPageGeometry,
    request_options: &docx_layout::types::LayoutOptions,
    regions: &DocumentRegions,
    pages: usize,
    anchored: bool,
) -> Result<(Vec<BlockExtent>, Vec<FontChainDependencies>), String> {
    let mut measures = Vec::new();
    let mut dependencies = Vec::new();
    let mut step = 32;
    loop {
        let start = measures.len();
        let end = prefix_boundary(blocks, (start + step).min(blocks.len()));
        if anchored {
            let mut prefix = blocks[..end].to_vec();
            let mut flow = docx_layout::measure_blocks::FloatFlow::with_table_wrap_frames(
                &prefix,
                &widths[..end],
                &table_wrap_frames[..end],
                measurement,
                Some(geometry),
            )?;
            flow.measure_until(&mut prefix, &widths[..end], measurement, end)?;
            dependencies = flow.font_dependencies().to_vec();
            measures = flow.into_extents();
            for (block, measured) in blocks.iter_mut().zip(prefix) {
                *block = measured;
            }
        } else {
            let mut flow = docx_layout::measure_blocks::FloatFlow::with_table_wrap_frames(
                &blocks[start..end],
                &widths[start..end],
                &table_wrap_frames[start..end],
                measurement,
                Some(geometry),
            )?;
            flow.measure_until(
                &mut blocks[start..end],
                &widths[start..end],
                measurement,
                end - start,
            )?;
            dependencies.extend_from_slice(flow.font_dependencies());
            measures.extend(flow.into_extents());
        }
        if end == blocks.len() {
            return Ok((measures, dependencies));
        }
        let options =
            options_through_section(request_options, regions, section_breaks(&blocks[..end]));
        // Columns are balanced over their whole section, so no prefix ends inside one.
        if options
            .columns
            .as_ref()
            .is_some_and(|columns| columns.count > 1.0)
        {
            step *= 2;
            continue;
        }
        let mut probe = LayoutInput {
            measured: blocks[..end]
                .iter()
                .zip(&measures)
                .map(|(block, measure)| MeasuredBlock {
                    block: block.clone(),
                    measure: measure.clone(),
                })
                .collect(),
            options,
        };
        let probed =
            docx_layout::place::layout_document(&mut probe).map_err(layout_error_message)?;
        if probed.pages.len() >= pages + 2 {
            return Ok((measures, dependencies));
        }
        step *= 2;
    }
}

/// A wrapped shape whose horizontal position depends on the page it lands on.
fn wraps_by_page_side(shape: &docx_layout::types::ShapeBlock) -> bool {
    let Some(horizontal) = shape
        .position
        .as_ref()
        .and_then(|position| position.horizontal.as_ref())
    else {
        return false;
    };
    matches!(
        shape.wrap_type.as_deref(),
        Some("square" | "tight" | "through" | "topAndBottom")
    ) && (horizontal.align.as_deref() == Some("inside")
        || matches!(
            horizontal.relative_to.as_deref(),
            Some("insideMargin" | "outsideMargin")
        ))
}

fn has_wrap_stabilized_shapes(blocks: &[LayoutBlock]) -> bool {
    blocks
        .iter()
        .any(|block| matches!(block, LayoutBlock::Shape(shape) if wraps_by_page_side(shape)))
}

/// What each page's note areas show, which an edit elsewhere can change.
fn note_page_keys(layout: Option<&Layout>) -> Vec<Option<Vec<NoteAreaContract>>> {
    layout.map_or_else(Vec::new, |layout| {
        layout
            .pages
            .iter()
            .map(|page| page.note_areas.clone())
            .collect()
    })
}

fn float_geometry_key(geometry: &docx_layout::measure_blocks::FloatPageGeometry) -> [f64; 5] {
    [
        geometry.page_width,
        geometry.margin_left,
        geometry.page_height,
        geometry.margin_top,
        geometry.content_height,
    ]
}

/// Whether any block holds an object that may float, whether or not it forms
/// a zone at the body's first width.
fn anchors_objects(blocks: &[LayoutBlock]) -> bool {
    blocks.iter().any(|block| match block {
        LayoutBlock::Paragraph(paragraph) => paragraph.runs.iter().any(|run| {
            matches!(
                run,
                docx_layout::types::Run::Image(image)
                    if image.position.is_some()
                        || image.wrap_type.is_some()
                        || image.display_mode.as_deref() == Some("float")
            )
        }),
        LayoutBlock::Table(table) => table.floating.is_some(),
        LayoutBlock::TextBox(_) | LayoutBlock::Shape(_) => true,
        _ => false,
    })
}

/// Whether every float in `blocks` hangs from the text before it, so a float
/// past a prefix cannot reach back into its pages, and no float waits on the
/// page-side wrapping that settles only after the whole body is placed.
fn floats_follow_the_text(blocks: &[LayoutBlock]) -> bool {
    fn text_relative(position: Option<&docx_layout::types::ImageRunPosition>) -> bool {
        position.is_none_or(|position| {
            position
                .vertical
                .as_ref()
                .and_then(|vertical| vertical.relative_to.as_deref())
                .is_some_and(|relative_to| matches!(relative_to, "paragraph" | "line"))
        })
    }
    blocks.iter().all(|block| match block {
        LayoutBlock::Paragraph(paragraph) => paragraph.runs.iter().all(|run| match run {
            docx_layout::types::Run::Image(image) => text_relative(image.position.as_ref()),
            _ => true,
        }),
        LayoutBlock::Table(table) => table
            .floating
            .as_ref()
            .is_none_or(|floating| floating.vert_anchor.as_deref() == Some("text")),
        LayoutBlock::TextBox(text_box) => text_relative(text_box.position.as_ref()),
        LayoutBlock::Shape(shape) => {
            text_relative(shape.position.as_ref()) && !wraps_by_page_side(shape)
        }
        _ => true,
    })
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
        .filter_map(|(index, measured)| match &measured.block {
            LayoutBlock::Shape(shape) if wraps_by_page_side(shape) => {
                Some((index, shape.id.clone()))
            }
            _ => None,
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
    let (widths, table_wrap_frames) = region_measurement_frames(blocks.iter(), input, regions);
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
        let measures = docx_layout::measure_blocks::measure_blocks_with_table_wrap_frames(
            &mut blocks,
            &widths,
            &table_wrap_frames,
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
            let variant = |kind: HeaderFooterKind, hf_type: HeaderFooterType| {
                variants.iter().rfind(|variant| {
                    variant.section_index == section_index
                        && variant.kind == kind
                        && variant.hf_type == hf_type
                })
            };
            let extend = |hf_type: HeaderFooterType| {
                let header = variant(HeaderFooterKind::Header, hf_type)
                    .map_or(0.0, |variant| variant.flow_height);
                let footer = variant(HeaderFooterKind::Footer, hf_type)
                    .map_or(0.0, |variant| variant.flow_height);
                extend_body_margins(&page_size, &margins, header, footer)
            };
            let float_bands = |hf_type: HeaderFooterType| {
                let mut bands: Vec<_> = [HeaderFooterKind::Header, HeaderFooterKind::Footer]
                    .into_iter()
                    .filter_map(|kind| variant(kind, hf_type))
                    .flat_map(|variant| {
                        header_footer_float_bands(
                            variant,
                            HeaderFooterMetrics {
                                kind: variant.kind,
                                page_size: &page_size,
                                margins: &margins,
                            },
                        )
                    })
                    .collect();
                bands.sort_by(|a, b| a.top.total_cmp(&b.top));
                bands
            };
            let even_and_odd =
                regions.even_and_odd_headers || section.even_and_odd_headers == Some(true);
            (
                extend(HeaderFooterType::Default),
                SectionPageMargins {
                    first: section.title_pg.then(|| extend(HeaderFooterType::First)),
                    even: even_and_odd.then(|| extend(HeaderFooterType::Even)),
                    restart: section
                        .page_numbering
                        .as_ref()
                        .and_then(|numbering| numbering.start),
                },
                docx_layout::types::SectionPageFloatBands {
                    default: float_bands(HeaderFooterType::Default),
                    first: section
                        .title_pg
                        .then(|| float_bands(HeaderFooterType::First)),
                    even: even_and_odd.then(|| float_bands(HeaderFooterType::Even)),
                    anchor_margins: Some(margins.clone()),
                },
            )
        })
        .collect();
    let (extended, bands): (Vec<_>, Vec<_>) = extended
        .into_iter()
        .map(|(margins, page_margins, float_bands)| (margins, (page_margins, float_bands)))
        .unzip();
    let (page_margins, float_bands): (Vec<_>, Vec<_>) = bands.into_iter().unzip();
    input.options.section_page_margins = page_margins
        .iter()
        .any(|margins| *margins != SectionPageMargins::default())
        .then_some(page_margins);
    input.options.section_page_float_bands = float_bands
        .iter()
        .any(|bands| {
            !bands.default.is_empty()
                || bands.first.as_ref().is_some_and(|bands| !bands.is_empty())
                || bands.even.as_ref().is_some_and(|bands| !bands.is_empty())
        })
        .then_some(float_bands);
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
    /// Non-font measurement config and font-store identity of `input`.
    measured_with: Option<u64>,
    /// The body lowering a float document's `input` arena was measured from.
    lowered_from: Option<Rc<Vec<LayoutBlock>>>,
    /// The doc epoch and lowering map a resident edit path built `input` from.
    input_lowering: Option<(u64, Rc<LoweringMap>)>,
    /// The widths, table frames and float geometry the arena was measured at,
    /// and whether floating zones shaped it.
    measured_widths: Vec<f64>,
    measured_font_dependencies: Vec<FontChainDependencies>,
    measured_table_wrap_frames: Vec<bool>,
    measured_float_geometry: Option<[f64; 5]>,
    measured_with_floats: bool,
    /// Pages whose note areas the last region pass changed.
    note_changed_pages: Vec<usize>,
    layout: Option<Layout>,
    checkpoints: Vec<LayoutCheckpoint>,
    block_fingerprints: Vec<Fingerprint>,
    options_fingerprint: u64,
    revision_preview_key: u64,
    /// The decisions [`Self::revision_preview_key`] identifies.
    revision_preview: BTreeMap<String, RevisionPreview>,
    /// The document epoch the retained input was lowered at.
    doc_epoch: u64,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    /// The pages the last pass placed afresh, within the range above.
    rebuilt_page_ranges: Vec<std::ops::Range<usize>>,
    position_deltas: HashMap<String, i64>,
    last_incremental: bool,
    /// Pages whose stamps changed since the last display build, when known.
    restamped_pages: Option<BTreeSet<usize>>,
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
    encoded_doc_epoch: u64,
    encoded_layout_epoch: u64,
    pages: Vec<FramePageSnapshot>,
    next_page_id: u64,
    font_cache_identity: Option<(u64, u64)>,
    font_chains: BTreeMap<String, Vec<u32>>,
    extras_fingerprint: u64,
    extras_json: Option<String>,
    /// The next frame is full whatever epoch the caller holds.
    fresh_base: bool,
    incremental_display_builds: u64,
    rebuilt_display_pages: u64,
    window: Option<std::ops::Range<usize>>,
    retain_built_pages: bool,
    windowed_incremental_builds: bool,
}

/// Engine observability snapshot.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EngineStats {
    pub doc_epoch: u64,
    pub lowered_story_count: usize,
    pub lowered_block_count: usize,
    pub lower_cache_hits: u64,
    pub lower_cache_misses: u64,
    pub lower_preview_patches: u64,
    pub lower_preview_fallbacks: u64,
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
/// The yrs transaction observer advances `doc_epoch` for every store-changing
/// local or remote commit. Render caches are generation-tagged instead of being
/// eagerly cleared, so an in-flight read can never publish blocks from a
/// different document generation.
pub struct EngineSession {
    doc: EditingDoc,
    doc_epoch: Rc<Cell<u64>>,
    // Kept alive for the lifetime of the document. Dropping it unregisters the
    // observer before the Rc epoch source is released.
    _doc_epoch_observer: Subscription,
    render: RefCell<RenderState>,
    note_separators: RefCell<Option<LoweredNoteSeparators>>,
    preview_font_requirements: RefCell<Option<PreviewFontRequirements>>,
    measurement: RefCell<MeasurementState>,
    regions: RefCell<Option<ResidentRegionState>>,
    pagination: RefCell<PaginationState>,
    display: RefCell<DisplayState>,
    /// The resident caret head and the body or nested body story it lies in.
    resident_caret_head: RefCell<Option<(String, StickyIndex)>>,
    capture: RefCell<Option<LayoutCapture>>,
    /// A region layout measured a step at a time, between two of its steps.
    resumable: RefCell<Option<ResumableRegionLayout>>,
    /// Content fingerprints of measurement fonts, by font store and font id.
    font_fingerprints: RefCell<HashMap<(u64, u32), String>>,
    /// The document holds part of a package, such as a preview's first blocks.
    partial_document: Cell<bool>,
    /// Resident text edits re-lower only their paragraph when eligible.
    local_lowering: Cell<bool>,
}

/// The font requirements of `blocks` that `measurement` gives no chain of registered fonts, so
/// they were measured with synthetic metrics.
fn missing_font_chains<'a>(
    measurement: &serde_json::Value,
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
) -> Vec<String> {
    let chains = measurement.get("fontChains");
    let defaults = measurement.get("defaults").cloned().unwrap_or_default();
    docx_layout::measure_blocks::collect_font_requirements(
        blocks,
        docx_layout::measure_blocks::default_font_family(&defaults),
    )
    .into_iter()
    .filter(|requirement| {
        !chains
            .and_then(|chains| chains.get(&requirement.key))
            .and_then(serde_json::Value::as_array)
            .is_some_and(|ids| {
                !ids.is_empty()
                    && ids.iter().all(|id| {
                        id.as_u64()
                            .and_then(|id| u32::try_from(id).ok())
                            .and_then(|id| docx_layout::with_measure_face(id, |_, _| ()))
                            .is_some()
                    })
            })
    })
    .map(|requirement| requirement.key)
    .collect()
}

/// A fingerprint of everything in a region layout request that shapes pages: sections,
/// settings, notes, the render environment and options, but not the fonts, which are
/// fingerprinted by content, or the gap between pages.
fn layout_options_fingerprint(mut request: serde_json::Value) -> String {
    if let Some(fields) = request.as_object_mut() {
        fields.remove("measurement");
        fields.remove("measured");
        if let Some(options) = fields
            .get_mut("options")
            .and_then(serde_json::Value::as_object_mut)
        {
            options.remove("pageGap");
        }
    }
    pages::sha256_hex(canonical_json(&request).as_bytes())
}

/// Identifies the revision preview a pass lays out under; 0 is none.
fn revision_preview_key(env: &RenderEnv) -> Result<u64, String> {
    if env.revision_preview.is_empty() {
        return Ok(0);
    }
    serde_json::to_vec(&env.revision_preview)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint revision preview: {error}"))
}

fn font_requirements_fingerprint(mut request: serde_json::Value) -> Result<u64, String> {
    if let Some(env) = request
        .get_mut("renderEnv")
        .and_then(serde_json::Value::as_object_mut)
    {
        env.remove("revisionPreview");
    }
    serde_json::to_vec(&request)
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint font requirements: {error}"))
}

/// `value` as JSON with object keys sorted, whatever order they were read in.
fn canonical_json(value: &serde_json::Value) -> String {
    match value {
        serde_json::Value::Object(fields) => {
            let sorted: BTreeMap<&String, String> = fields
                .iter()
                .map(|(key, value)| (key, canonical_json(value)))
                .collect();
            let body = sorted
                .into_iter()
                .map(|(key, value)| format!("{}:{value}", serde_json::Value::from(key.as_str())))
                .collect::<Vec<_>>()
                .join(",");
            format!("{{{body}}}")
        }
        serde_json::Value::Array(items) => format!(
            "[{}]",
            items
                .iter()
                .map(canonical_json)
                .collect::<Vec<_>>()
                .join(",")
        ),
        other => other.to_string(),
    }
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// Whether a display page shows the section and numbering stamps of `page`,
/// which choose its header and footer and fill its page fields.
fn page_stamps_match(
    shown: &docx_layout::display_list::DisplayPage,
    page: &docx_layout::types::Page,
) -> bool {
    shown.section_id == page.section_id
        && shown.section_index == page.section_index
        && shown.section_page_index == page.section_page_index
        && shown.section_page_number == page.section_page_number
        && shown.page_label == page.page_label
}

/// Which pages a full display build compiles: all of them without a window,
/// otherwise the window plus every page the previous list had built.
fn full_build_pages(display: &DisplayState, page_count: usize) -> Option<BTreeSet<usize>> {
    let window = display.window.as_ref()?;
    let mut pages: BTreeSet<_> = (window.start..window.end.min(page_count)).collect();
    if let Some(list) = &display.list {
        pages.extend(
            list.pages
                .iter()
                .enumerate()
                .take(page_count)
                .filter_map(|(index, page)| (!page.unbuilt).then_some(index)),
        );
    }
    Some(pages)
}

/// A whole-document placement pass, as one rebuilt page range.
fn full_pass(input: &mut LayoutInput) -> Result<docx_layout::place::IncrementalLayout, String> {
    let run =
        docx_layout::place::layout_document_checkpointed(input).map_err(|error| match error {
            docx_layout::LayoutError::Unsupported(_) => "UNSUPPORTED".to_owned(),
            docx_layout::LayoutError::Invalid(reason) => reason,
        })?;
    Ok(docx_layout::place::IncrementalLayout {
        rebuilt_page_ranges: std::iter::once(run.rebuilt_page_start..run.rebuilt_page_end)
            .collect(),
        checkpointed: run,
    })
}

/// Where the resident caret sits, for the pages a windowed build keeps.
#[derive(Clone, Copy)]
enum CaretExtent {
    /// Its display position, through the lowering the layout was built from.
    Position(f64),
    /// No exact position (inside an atom or hidden text, or no retained
    /// lowering), so every re-placed page is kept.
    Unmapped,
}

impl CaretExtent {
    /// Whether a fragment over display positions `start..=end` holds the caret.
    fn holds(self, start: f64, end: f64) -> bool {
        match self {
            Self::Position(position) => start <= position && position <= end,
            Self::Unmapped => true,
        }
    }
}

/// Selects the window and caret pages among the rebuilt pages.
fn window_build_pages(
    display: &DisplayState,
    layout: &Layout,
    rebuilt_pages: impl IntoIterator<Item = usize>,
    caret: Option<CaretExtent>,
) -> Option<BTreeSet<usize>> {
    if !display.windowed_incremental_builds
        || display.window.is_none()
        || display.retain_built_pages
    {
        return full_build_pages(display, layout.pages.len());
    }
    let window = display.window.as_ref()?;
    let mut pages: BTreeSet<_> = (window.start..window.end.min(layout.pages.len())).collect();
    if let Some(caret) = caret {
        for index in rebuilt_pages {
            let Some(page) = layout.pages.get(index) else {
                continue;
            };
            if pages.contains(&index) {
                continue;
            }
            if matches!(caret, CaretExtent::Unmapped)
                || page.fragments.iter().any(|fragment| {
                    let (start, end) = match fragment {
                        Fragment::Paragraph(value) => (value.pm_start, value.pm_end),
                        Fragment::Table(value) => (value.pm_start, value.pm_end),
                        Fragment::Image(value) => (value.pm_start, value.pm_end),
                        Fragment::Shape(value) => (value.pm_start, value.pm_end),
                        Fragment::Chart(value) => (value.pm_start, value.pm_end),
                        Fragment::TextBox(value) => (value.pm_start, value.pm_end),
                    };
                    start
                        .zip(end)
                        .is_some_and(|(start, end)| caret.holds(start, end))
                })
            {
                pages.insert(index);
            }
        }
    }
    Some(pages)
}

/// With windowed builds on, compile the window and a mapped caret's page,
/// or every built page while retained; otherwise use [`full_build_pages`].
fn windowed_full_build_pages(
    display: &DisplayState,
    layout: &Layout,
    caret: Option<CaretExtent>,
) -> Option<BTreeSet<usize>> {
    window_build_pages(
        display,
        layout,
        0..layout.pages.len(),
        caret.filter(|caret| matches!(caret, CaretExtent::Position(_))),
    )
}

fn resident_paragraph_position(input: &LayoutInput, para_id: &str, offset: u32) -> Option<i64> {
    let start = input.measured.iter().find_map(|measured| {
        let (id, start) = paragraph_identity(&measured.block)?;
        (block_key(id) == para_id).then_some(start?)
    })?;
    paragraph_offset_position(start, offset)
}

/// The display positions every block lowered from the paragraph `para_id` spans,
/// in table cells too.
fn resident_paragraph_span(input: &LayoutInput, para_id: &str) -> Option<(f64, f64)> {
    fn visit(block: &LayoutBlock, para_id: &str, span: &mut Option<(f64, f64)>) {
        match block {
            LayoutBlock::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .flat_map(|cell| &cell.blocks)
                .for_each(|nested| visit(nested, para_id, span)),
            LayoutBlock::Paragraph(paragraph) if block_key(&paragraph.id) == para_id => {
                if let (Some(start), Some(end)) = (paragraph.pm_start, paragraph.pm_end) {
                    let (low, high) = span.get_or_insert((start, end));
                    *low = low.min(start);
                    *high = high.max(end);
                }
            }
            _ => {}
        }
    }
    let mut span = None;
    for measured in &input.measured {
        visit(&measured.block, para_id, &mut span);
    }
    span
}

/// The display position of story index `index` in `paragraph` of `story`, through the lowering
/// `input` was built from. `None` inside an atom, expanded inline content or hidden text, at a
/// boundary its spans disagree on, or when the lowering does not start the paragraph where
/// `input` does.
fn lowered_caret_position(
    map: &LoweringMap,
    input: &LayoutInput,
    story: &str,
    paragraph: &crate::segments::ParaEntry,
    index: u32,
) -> Option<f64> {
    let source = map.paragraphs.iter().position(|(slot, id)| {
        id.as_str() == &*paragraph.para_id
            && map
                .stories
                .get(*slot as usize)
                .is_some_and(|name| name == story)
    })?;
    let source = u32::try_from(source).ok()?;
    let first_block = map
        .paragraph_blocks
        .iter()
        .filter(|(_, block)| *block == source)
        .map(|(pm, _)| *pm)
        .min()?;
    let (start, _) = resident_paragraph_span(input, &paragraph.para_id)?;
    if first_block as f64 != start {
        return None;
    }
    let mut spans = map
        .spans
        .iter()
        .filter(|span| span.paragraph == source)
        .peekable();
    if spans.peek().is_none() {
        return (index == paragraph.node_start).then_some(start + 1.0);
    }
    span_position(spans, index).map(|position| position as f64)
}

/// The display position of story index `index` that every span touching it agrees on.
fn span_position<'a>(
    spans: impl Iterator<Item = &'a crate::bridge::SourceSpan>,
    index: u32,
) -> Option<u64> {
    let mut position = None;
    for span in spans.filter(|span| (span.raw_start..=span.raw_end).contains(&index)) {
        let units = u64::from(span.raw_end.checked_sub(span.raw_start)?);
        let candidate = if !span.atom && span.pm_end.checked_sub(span.pm_start) == Some(units) {
            span.pm_start + u64::from(index - span.raw_start)
        } else if index == span.raw_start {
            span.pm_start
        } else if index == span.raw_end {
            span.pm_end
        } else {
            return None;
        };
        if position.is_some_and(|position| position != candidate) {
            return None;
        }
        position = Some(candidate);
    }
    position
}

fn paragraph_offset_position(start: f64, offset: u32) -> Option<i64> {
    (start.is_finite()
        && start.fract() == 0.0
        && start >= i64::MIN as f64
        && start <= i64::MAX as f64)
        .then_some(start as i64)?
        .checked_add(1 + i64::from(offset))
}

fn measured_fingerprint(measured: &MeasuredBlock) -> Result<Fingerprint, String> {
    crate::fingerprint::fingerprint_without_positions(&(
        measured,
        relative_run_position_fingerprint(&measured.block),
    ))
    .map_err(|error| format!("fingerprint measured block: {error}"))
}

fn relative_run_position_fingerprint(block: &LayoutBlock) -> u64 {
    fn paragraph(block: &docx_layout::types::ParagraphBlock, positions: &mut DefaultHasher) {
        let start = block.pm_start.unwrap_or(0.0);
        for (from, to) in std::iter::once((Some(0.0), block.pm_end.map(|end| end - start))).chain(
            block.runs.iter().map(|run| {
                (
                    run.pm_start().map(|position| position - start),
                    run.pm_end().map(|position| position - start),
                )
            }),
        ) {
            for position in [from, to] {
                positions.write_u8(u8::from(position.is_some()));
                if let Some(position) = position {
                    positions.write_u64(if position == 0.0 {
                        0
                    } else {
                        position.to_bits()
                    });
                }
            }
        }
    }

    fn collect(block: &LayoutBlock, positions: &mut DefaultHasher) {
        match block {
            LayoutBlock::Paragraph(block) => paragraph(block, positions),
            LayoutBlock::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        for block in &cell.blocks {
                            collect(block, positions);
                        }
                    }
                }
            }
            LayoutBlock::TextBox(textbox) => {
                for block in &textbox.content {
                    paragraph(block, positions);
                }
            }
            LayoutBlock::Shape(shape) => {
                for block in shape.inner_text.iter().flatten().chain(
                    shape
                        .children
                        .iter()
                        .flat_map(|child| child.inner_text.iter().flatten()),
                ) {
                    paragraph(block, positions);
                }
            }
            _ => {}
        }
    }

    let mut positions = DefaultHasher::new();
    collect(block, &mut positions);
    positions.finish()
}

/// [`measured_fingerprint`] of a block and its measure held apart.
fn measured_parts_fingerprint(
    block: &LayoutBlock,
    measure: &BlockExtent,
) -> Result<Fingerprint, String> {
    #[derive(Serialize)]
    struct MeasuredParts<'a> {
        block: &'a LayoutBlock,
        measure: &'a BlockExtent,
    }
    crate::fingerprint::fingerprint_without_positions(&(
        MeasuredParts { block, measure },
        relative_run_position_fingerprint(block),
    ))
    .map_err(|error| format!("fingerprint measured block: {error}"))
}

/// JSON equality with numbers compared by value, as a host's `1` and Rust's `1.0`.
fn json_equal(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    use serde_json::Value;
    match (left, right) {
        (Value::Number(left), Value::Number(right)) => {
            left == right
                || left
                    .as_f64()
                    .is_some_and(|value| Some(value) == right.as_f64())
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len() && left.iter().zip(right).all(|(l, r)| json_equal(l, r))
        }
        (Value::Object(left), Value::Object(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .all(|(key, l)| right.get(key).is_some_and(|r| json_equal(l, r)))
        }
        _ => left == right,
    }
}

fn json_option_equal(left: Option<&serde_json::Value>, right: Option<&serde_json::Value>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => json_equal(left, right),
        (left, right) => left.is_none() && right.is_none(),
    }
}

/// A section break's extent never depends on its margins, which header and
/// footer extents widen after measurement.
fn section_breaks_match_but_margins(next: &LayoutBlock, retained: &LayoutBlock) -> bool {
    let (LayoutBlock::SectionBreak(next), LayoutBlock::SectionBreak(retained)) = (next, retained)
    else {
        return false;
    };
    let mut retained = retained.clone();
    retained.margins.clone_from(&next.margins);
    *next == retained
}

/// `dirty`, or the start of the section whose closing break changed first, if
/// that is earlier: a section break carries geometry (margins a header
/// widens, page size, columns) that the whole section was laid out under, so
/// pagination has to resume where the section begins.
fn section_start_of_first_changed_break(
    measured: &[MeasuredBlock],
    previous: &[Fingerprint],
    next: &[Fingerprint],
    dirty: usize,
) -> usize {
    let Some(changed) = (dirty..measured.len()).find(|&index| {
        matches!(measured[index].block, LayoutBlock::SectionBreak(_))
            && previous.get(index) != next.get(index)
    }) else {
        return dirty;
    };
    // From the break that opens the section: the section's first page starts there.
    let section_start = measured[..changed]
        .iter()
        .rposition(|block| matches!(block.block, LayoutBlock::SectionBreak(_)))
        .unwrap_or(0);
    dirty.min(section_start)
}

fn measured_fingerprints(input: &LayoutInput) -> Result<Vec<Fingerprint>, String> {
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

/// A table cell or content control story lowered as part of the body.
fn is_nested_body_story(story: &str) -> bool {
    story.starts_with("body:")
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

type ResidentWalkOut<'a> = (
    &'a mut Vec<MeasuredBlock>,
    &'a mut Vec<Fingerprint>,
    &'a mut Vec<(usize, usize)>,
    &'a mut Vec<FontChainDependencies>,
);

/// The dirty-block walk of `resident_layout_input_from_blocks`. With `take`, reused extents move
/// out of `previous` and `moved` records `(previous index, measured index)` for each one.
fn resident_walk(
    blocks: &[LayoutBlock],
    any_block: bool,
    paragraph_merge: bool,
    (previous, previous_fingerprints, previous_dependencies, take): (
        &mut [MeasuredBlock],
        &[Fingerprint],
        &[FontChainDependencies],
        bool,
    ),
    measure_dirty: &mut dyn FnMut(
        usize,
        &str,
        &LayoutBlock,
        &mut LayoutBlock,
    ) -> Result<BlockExtent, String>,
    (measured, block_fingerprints, moved, dependencies): ResidentWalkOut<'_>,
) -> Result<(u64, u64), String> {
    let structure = || "resident plain-text input changed the block structure".to_owned();
    let identity = || "resident plain-text input changed stable block identity".to_owned();
    let mut reuse = |previous: &mut [MeasuredBlock], index: usize, to: usize| {
        if take {
            moved.push((index, to));
            std::mem::replace(&mut previous[index].measure, BlockExtent::Unsupported)
        } else {
            previous[index].measure.clone()
        }
    };
    let mut cursor = 0;
    let mut skipped_merged_paragraph = false;
    let mut resident_measure_calls = 0_u64;
    let mut resident_reused_blocks = 0_u64;
    for (block_index, next_block) in blocks.iter().enumerate() {
        let mut index = cursor;
        if index >= previous.len() {
            return Err(structure());
        }
        cursor += 1;
        if paragraph_merge && !resident_block_slots_match(&previous[index].block, next_block) {
            if skipped_merged_paragraph || paragraph_identity(&previous[index].block).is_none() {
                return Err(structure());
            }
            skipped_merged_paragraph = true;
            index = cursor;
            if index >= previous.len() {
                return Err(structure());
            }
            cursor += 1;
        }
        if paragraph_merge && !resident_block_slots_match(&previous[index].block, next_block) {
            return Err(identity());
        }
        let previous_fingerprint = previous_fingerprints[index];
        let previous_block = &previous[index].block;
        let (Some((next_id, _)), Some((previous_id, _))) = (
            paragraph_identity(next_block),
            paragraph_identity(previous_block),
        ) else {
            if next_block != previous_block {
                let (true, Some(next_id), Some(previous_id)) = (
                    any_block,
                    fragment_identity(next_block),
                    fragment_identity(previous_block),
                ) else {
                    return Err(
                        "resident plain-text input changed a non-paragraph block".to_owned()
                    );
                };
                let key = block_key(next_id);
                if key != block_key(previous_id) {
                    return Err(identity());
                }
                let mut next_measured_block = next_block.clone();
                let (measure, reads) = FontChainDependencies::capture(|| {
                    measure_dirty(block_index, &key, previous_block, &mut next_measured_block)
                });
                let measure = measure?;
                dependencies.push(reads);
                let measured_block = MeasuredBlock {
                    block: next_measured_block,
                    measure,
                };
                block_fingerprints.push(measured_fingerprint(&measured_block)?);
                measured.push(measured_block);
                resident_measure_calls = resident_measure_calls.wrapping_add(1);
                continue;
            }
            dependencies.push(
                previous_dependencies
                    .get(index)
                    .cloned()
                    .unwrap_or_else(FontChainDependencies::unknown),
            );
            let measure = reuse(previous, index, measured.len());
            measured.push(MeasuredBlock {
                block: next_block.clone(),
                measure,
            });
            block_fingerprints.push(previous_fingerprint);
            resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
            continue;
        };
        let key = block_key(next_id);
        if key != block_key(previous_id) {
            return Err(identity());
        }
        if next_block == previous_block {
            dependencies.push(
                previous_dependencies
                    .get(index)
                    .cloned()
                    .unwrap_or_else(FontChainDependencies::unknown),
            );
            let measure = reuse(previous, index, measured.len());
            measured.push(MeasuredBlock {
                block: next_block.clone(),
                measure,
            });
            block_fingerprints.push(previous_fingerprint);
            resident_reused_blocks = resident_reused_blocks.wrapping_add(1);
            continue;
        }

        let mut next_measured_block = next_block.clone();
        let (measure, reads) = FontChainDependencies::capture(|| {
            measure_dirty(block_index, &key, previous_block, &mut next_measured_block)
        });
        let measure = measure?;
        dependencies.push(reads);
        let measured_block = MeasuredBlock {
            block: next_measured_block,
            measure,
        };
        block_fingerprints.push(measured_fingerprint(&measured_block)?);
        measured.push(measured_block);
        resident_measure_calls = resident_measure_calls.wrapping_add(1);
    }
    if cursor < previous.len() {
        if !paragraph_merge
            || skipped_merged_paragraph
            || paragraph_identity(&previous[cursor].block).is_none()
            || cursor + 1 < previous.len()
        {
            return Err(structure());
        }
        skipped_merged_paragraph = true;
    }
    if paragraph_merge && !skipped_merged_paragraph {
        return Err(structure());
    }
    Ok((resident_measure_calls, resident_reused_blocks))
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

/// Whether `block` is, or nests, one of `paragraphs`.
fn block_holds_paragraph(block: &LayoutBlock, paragraphs: &HashSet<String>) -> bool {
    if let LayoutBlock::Paragraph(paragraph) = block {
        return paragraphs.contains(block_key(&paragraph.id).as_ref());
    }
    let mut keys = Vec::new();
    nested_block_keys(block, &mut keys);
    keys.iter().any(|key| paragraphs.contains(key))
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
            .observe_after_transaction(move |txn| {
                if !txn.delete_set().is_empty() || txn.after_state() != txn.before_state() {
                    observer_epoch.set(observer_epoch.get().wrapping_add(1));
                }
            })
            .expect("EngineSession document transaction observer registers");
        Self {
            doc,
            doc_epoch,
            _doc_epoch_observer: observer,
            render: RefCell::new(RenderState::default()),
            note_separators: RefCell::new(None),
            preview_font_requirements: RefCell::new(None),
            measurement: RefCell::new(MeasurementState::default()),
            regions: RefCell::new(None),
            pagination: RefCell::new(PaginationState::default()),
            display: RefCell::new(DisplayState::default()),
            resident_caret_head: RefCell::new(None),
            capture: RefCell::new(None),
            resumable: RefCell::new(None),
            font_fingerprints: RefCell::new(HashMap::new()),
            partial_document: Cell::new(false),
            local_lowering: Cell::new(false),
        }
    }

    /// Lets an eligible resident text edit re-lower only its paragraph. Off by default.
    pub fn set_local_lowering(&self, enabled: bool) {
        self.local_lowering.set(enabled);
    }

    /// Marks whether the document is part of a package, such as a preview's
    /// first blocks: such a document's layouts count only its own pages, so
    /// they render NUMPAGES empty.
    pub fn set_partial_document(&self, partial: bool) {
        self.partial_document.set(partial);
    }

    /// Editing document.
    pub fn doc(&self) -> &EditingDoc {
        &self.doc
    }

    pub fn doc_epoch(&self) -> u64 {
        self.doc_epoch.get()
    }

    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn edit_resident_text(
        &self,
        range: crate::StoryRange,
        text: Option<&str>,
        lower_locally: bool,
    ) -> crate::OpResult<crate::Receipt> {
        let before = self.doc_epoch();
        let index_epoch = (self.local_lowering.get()
            && (text.is_none() || range.start == range.end))
            .then(|| self.doc.committed_epoch());
        let plain_text_delete = if index_epoch.is_some() && text.is_none() {
            let segments = self.doc.segment_index(&range.story)?;
            segments.is_text_range(range.start, range.end)
                && (segments
                    .segment_at(range.end)
                    .is_none_or(|segment| !matches!(segment.kind, crate::segments::SegKind::Embed))
                    || self
                        .doc
                        .paragraph_index(&range.story)?
                        .para_at(range.start)
                        .is_some_and(|paragraph| range.start > paragraph.node_start))
        } else {
            false
        };
        let length_before = if plain_text_delete || (index_epoch.is_some() && text.is_some()) {
            Some(self.doc.story_len(&range.story)?)
        } else {
            None
        };
        let mut attrs = None;
        let ctx = crate::EditCtx::local("", "");
        let receipt = match text {
            Some(text) => self.doc.insert_text_observed(
                &ctx,
                crate::Position::new(&range.story, range.start),
                text,
                crate::FormatPolicy::Inherit,
                |effective| attrs = Some(effective.clone()),
            )?,
            None => self.doc.delete_range(&ctx, range.clone())?,
        };
        if let Some(before) = index_epoch {
            match text {
                Some(text)
                    if length_before
                        .zip(self.doc.story_len(&range.story).ok())
                        .is_some_and(|(before, after)| {
                            after.checked_sub(before) == Some(text.encode_utf16().count() as u32)
                        }) =>
                {
                    self.doc.advance_indexes_after_text_insert(
                        &range.story,
                        before,
                        self.doc.committed_epoch(),
                        range.start,
                        text,
                    );
                }
                None if plain_text_delete
                    && receipt.new_para_ids.is_empty()
                    && receipt.revision_ids.is_empty()
                    && length_before
                        .zip(self.doc.story_len(&range.story).ok())
                        .is_some_and(|(before, after)| {
                            before.checked_sub(after) == Some(range.end - range.start)
                        }) =>
                {
                    self.doc.advance_indexes_after_text_delete(
                        &range.story,
                        before,
                        self.doc.committed_epoch(),
                        range.start,
                        range.end,
                    );
                }
                _ => {}
            }
        }
        let mut render = self.render.borrow_mut();
        let eligible = lower_locally
            && self.local_lowering.get()
            && (text.is_none() || range.start == range.end);
        if let Some(lowered) = render.stories.get_mut(&range.story) {
            lowered.local.edit = receipt
                .range
                .as_ref()
                .filter(|_| eligible && lowered.doc_epoch == before)
                .and_then(|range_result| {
                    let paragraph = &range_result.start.para;
                    Some(crate::bridge::local::TextEdit {
                        offset: lowered.local.offset(paragraph, range.start)?,
                        removed: range.end.checked_sub(range.start)?,
                        paragraph: paragraph.clone(),
                        text: text.unwrap_or_default().to_owned(),
                        attributes: attrs,
                        epochs: (before, self.doc_epoch()),
                    })
                });
        }
        Ok(receipt)
    }

    fn patch_lowered_body(&self, epoch: u64, env: &RenderEnv) -> Option<()> {
        let mut render = self.render.borrow_mut();
        let lowered = render.stories.get_mut("body")?;
        let edit = lowered.local.edit.take()?;
        if edit.epochs != (lowered.doc_epoch, epoch)
            || epoch != lowered.doc_epoch.wrapping_add(1)
            || lowered.env != *env
            || lowered.media != self.doc.media_sources()
            || !lowered.revealable_blocks.is_empty()
            || !lowered.local.matches_source(&self.doc)
            || self.regions.borrow().as_ref().is_some_and(|state| {
                state.fast_path.as_ref().is_none_or(|fast| {
                    !fast.notes_clear || fast.regions.sections.is_empty() || fast.render_env != *env
                })
            })
        {
            return None;
        }
        self.pagination.borrow_mut().input_lowering = None;
        let blocks = Rc::get_mut(&mut lowered.blocks)?;
        let map = Rc::get_mut(&mut lowered.map)?;
        let txn = self.doc.yrs_doc().transact();
        lowered.local.patch(blocks, map, &txn, env, &edit)?;
        lowered.doc_epoch = epoch;
        lowered.serialized_blocks = None;
        lowered.preview = None;
        Some(())
    }

    fn patch_preview_body(&self, epoch: u64, env: &RenderEnv) -> Option<()> {
        let mut render = self.render.borrow_mut();
        let lowered = render.stories.get_mut("body")?;
        if lowered.doc_epoch != epoch
            || lowered.media != self.doc.media_sources()
            || !lowered.local.matches_source(&self.doc)
            || lowered.env.revision_preview == env.revision_preview
        {
            return None;
        }
        let mut previous = lowered.env.clone();
        previous.revision_preview = env.revision_preview.clone();
        if previous != *env {
            return None;
        }
        let changed = lowered
            .env
            .revision_preview
            .keys()
            .chain(env.revision_preview.keys())
            .filter(|id| lowered.env.revision_preview.get(*id) != env.revision_preview.get(*id))
            .cloned()
            .collect();
        let patched = (|| {
            if !lowered.local.blocked {
                return None;
            }
            let units = lowered.preview.as_ref()?;
            if crate::bridge::preview::targets(units, &changed) {
                let replays = crate::bridge::preview::replay(
                    &self.doc,
                    env,
                    units,
                    &changed,
                    &lowered.blocks,
                )?;
                let mut units = units.as_ref().clone();
                crate::bridge::preview::splice(
                    replays,
                    Rc::make_mut(&mut lowered.blocks),
                    Rc::make_mut(&mut lowered.map),
                    Rc::make_mut(&mut lowered.revealable_blocks),
                    &mut units,
                );
                lowered.preview = Some(Rc::new(units));
                lowered.serialized_blocks = None;
            }
            lowered.env = env.clone();
            Some(())
        })();
        if patched.is_some() {
            render.preview_patches = render.preview_patches.wrapping_add(1);
        } else {
            render.preview_fallbacks = render.preview_fallbacks.wrapping_add(1);
        }
        patched
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
            .is_some_and(|cached| {
                cached.doc_epoch == epoch
                    && cached.env == *env
                    && cached.media == self.doc.media_sources()
                    && cached.local.matches_source(&self.doc)
            })
    }

    fn lower_story_into_cache(
        &self,
        story: &str,
        epoch: u64,
        env: &RenderEnv,
    ) -> Result<(), BridgeError> {
        let mut local = crate::bridge::local::LocalLowering::new(self.local_lowering.get());
        // Preview units serve later decisions, so a story's first lowering (the open) skips them.
        let record = self.render.borrow().stories.contains_key(story);
        let (blocks, map, revealable_blocks, preview) =
            crate::bridge::preview::lower_recorded(&self.doc, story, env, &mut local, record)?;
        let mut render = self.render.borrow_mut();
        render.cache_misses = render.cache_misses.wrapping_add(1);
        render.stories.insert(
            story.to_owned(),
            LoweredStory {
                doc_epoch: epoch,
                env: env.clone(),
                media: self.doc.media_sources(),
                blocks: Rc::new(blocks),
                map: Rc::new(map),
                revealable_blocks: Rc::new(revealable_blocks),
                serialized_blocks: None,
                local,
                preview: preview.map(Rc::new),
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
        self.with_lowered_story_mapped(story, env, after_lower, |blocks, _| read(blocks))
    }

    /// [`Self::with_lowered_story_observed`], also handing `read` the doc epoch and lowering
    /// map of the same lowering as the blocks.
    fn with_lowered_story_mapped<T>(
        &self,
        story: &str,
        env: &RenderEnv,
        after_lower: &mut dyn FnMut(),
        read: impl FnOnce(&[LayoutBlock], (u64, Rc<LoweringMap>)) -> T,
    ) -> Result<T, BridgeError> {
        let epoch = self.doc_epoch();
        if self.story_is_resident(story, epoch, env) {
            let mut render = self.render.borrow_mut();
            render.cache_hits = render.cache_hits.wrapping_add(1);
        } else if story != "body"
            || (self.patch_preview_body(epoch, env).is_none()
                && self.patch_lowered_body(epoch, env).is_none())
        {
            self.lower_story_into_cache(story, epoch, env)?;
        }
        let (blocks, map) = {
            let render = self.render.borrow();
            let lowered = render
                .stories
                .get(story)
                .expect("resident story exists after lowering");
            (Rc::clone(&lowered.blocks), Rc::clone(&lowered.map))
        };
        after_lower();
        Ok(read(&blocks, (epoch, map)))
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
            lower_preview_patches: render.preview_patches,
            lower_preview_fallbacks: render.preview_fallbacks,
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
                .rebuilt_page_ranges
                .iter()
                .map(ExactSizeIterator::len)
                .sum(),
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
            let block: LayoutBlock =
                serde_json::from_value(template.envelope.get("block")?.clone()).ok()?;
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
        self.layout_font_requirements(input_json, true)
    }

    fn layout_font_requirements(
        &self,
        input_json: &str,
        use_preview_superset: bool,
    ) -> Result<String, String> {
        let request: RegionLayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        let (input, regions, notes, measurement, render_env, body_story) = request.split();
        let cache_key = if use_preview_superset
            && !RenderEnv::parse_revision_preview(&render_env["revisionPreview"]).is_empty()
        {
            let request =
                serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
            let separators = self
                .doc
                .note_separator_state()?
                .map_or(0, |state| hash_bytes(&state));
            let fingerprint = hash_bytes(
                &[
                    font_requirements_fingerprint(request)?.to_le_bytes(),
                    separators.to_le_bytes(),
                ]
                .concat(),
            );
            let epoch = self.doc_epoch();
            let cached = self
                .preview_font_requirements
                .borrow()
                .as_ref()
                .filter(|cached| {
                    cached.doc_epoch == epoch && cached.request_fingerprint == fingerprint
                })
                .map(|cached| cached.json.clone());
            match cached {
                Some(Some(json)) => return Ok(json),
                Some(None) => return self.layout_font_requirements(input_json, false),
                None => {}
            }
            Some((epoch, fingerprint))
        } else {
            None
        };
        let default_family =
            docx_layout::measure_blocks::default_font_family(&measurement.defaults);
        let mut collector = docx_layout::measure_blocks::FontRequirementCollector::default();
        if cache_key.is_some() {
            collector.collect_preview(
                input.measured.iter().map(|measured| &measured.block),
                default_family,
            );
        } else {
            collector.collect(
                input.measured.iter().map(|measured| &measured.block),
                default_family,
            );
        }
        if let Some(body_story) = body_story {
            let mut render_env: RenderEnv = serde_json::from_value(render_env)
                .map_err(|error| format!("parse render environment: {error}"))?;
            if cache_key.is_some() {
                render_env.revision_preview.clear();
            }
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
            for (kind, name) in [
                (NoteKind::Footnote, "footnote"),
                (NoteKind::Endnote, "endnote"),
            ] {
                if notes
                    .contents
                    .iter()
                    .any(|content| content.note_kind == kind)
                    && let Some(blocks) = self.lower_note_separator(name, &render_env)?
                {
                    if cache_key.is_some() {
                        collector.collect_preview(blocks.iter(), default_family);
                        if let Some(revealable) =
                            self.note_separator_revealable(name, &render_env)?
                        {
                            collector.collect_preview(revealable.iter(), default_family);
                        }
                    } else {
                        collector.collect(blocks.iter(), default_family);
                    }
                }
            }
            let note_stories: BTreeMap<_, _> = notes
                .contents
                .into_iter()
                .map(|content| {
                    let prefix = match content.note_kind {
                        docx_layout::footnotes::NoteKind::Footnote => "fn",
                        docx_layout::footnotes::NoteKind::Endnote => "en",
                    };
                    (format!("{prefix}:{}", content.id), content)
                })
                .collect();
            stories.extend(note_stories.keys().cloned());
            for story in stories {
                self.with_lowered_story(&story, &render_env, |blocks| {
                    let blocks = if let Some(content) = note_stories.get(&story) {
                        let mut blocks = blocks.to_vec();
                        apply_note_presentation(
                            &mut blocks,
                            content.display_number.unwrap_or(1),
                            content.display_label.as_deref().unwrap_or("1"),
                        );
                        Cow::Owned(blocks)
                    } else {
                        Cow::Borrowed(blocks)
                    };
                    if cache_key.is_some() {
                        collector.collect_preview(blocks.iter(), default_family);
                    } else {
                        collector.collect(blocks.iter(), default_family);
                    }
                })
                .map_err(|error| error.to_string())?;
                if cache_key.is_some() {
                    let revealable = Rc::clone(
                        &self
                            .render
                            .borrow()
                            .stories
                            .get(&story)
                            .expect("resident story exists after lowering")
                            .revealable_blocks,
                    );
                    collector.collect_preview(revealable.iter(), default_family);
                }
            }
        }
        let requirements = collector.finish();
        let preview_superset_safe = requirements
            .values()
            .all(|requirement| requirement.scripts.is_empty());
        if let Some((doc_epoch, request_fingerprint)) = cache_key
            && !preview_superset_safe
        {
            self.preview_font_requirements
                .replace(Some(PreviewFontRequirements {
                    doc_epoch,
                    request_fingerprint,
                    json: None,
                }));
            return self.layout_font_requirements(input_json, false);
        }
        let json = serde_json::to_string(&requirements.into_values().collect::<Vec<_>>())
            .map_err(|error| format!("serialize: {error}"))?;
        if let Some((doc_epoch, request_fingerprint)) = cache_key {
            self.preview_font_requirements
                .replace(Some(PreviewFontRequirements {
                    doc_epoch,
                    request_fingerprint,
                    json: Some(json.clone()),
                }));
        }
        Ok(json)
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
        self.layout_regions_retained_json(input_json, None)
    }

    pub fn retained_layout_json(&self) -> Result<String, String> {
        let pass = {
            let regions = self.regions.borrow();
            let state = regions
                .as_ref()
                .ok_or("resident region layout is not built")?;
            RegionPass {
                notes_converged: state.notes_converged,
                provisional: state.provisional,
            }
        };
        self.retained_region_layout_json(&pass)
    }

    /// [`Self::layout_document_with_regions_retained_json`] for a caller that
    /// reads only the retained state: the layout is not serialized.
    pub fn layout_document_with_regions_retained(&self, input_json: &str) -> Result<(), String> {
        self.layout_document_with_regions_value(input_json)
            .map(|_| ())
    }

    /// The retained region layout's `headersFooters`, serialized as its
    /// retained reply carries them.
    pub fn retained_headers_footers_json(&self) -> Result<Option<String>, String> {
        self.regions
            .borrow()
            .as_ref()
            .and_then(|state| state.headers_footers.as_ref())
            .map(|value| {
                serde_json::to_string(value).map_err(|error| format!("serialize: {error}"))
            })
            .transpose()
    }

    /// [`Self::layout_document_with_regions_retained_json`] over only as much
    /// of the body as fills the first `pages` pages. A reply marked
    /// `provisional` holds a layout of that prefix: its page count is the
    /// prefix's, NUMPAGES fields render empty, and retained state serves only
    /// display and caret reads until a full pass replaces it. Edits made
    /// meanwhile take the full region pass.
    pub fn layout_document_with_regions_prefix_retained_json(
        &self,
        input_json: &str,
        pages: usize,
    ) -> Result<String, String> {
        self.layout_regions_retained_json(input_json, Some(pages))
    }

    /// [`Self::layout_document_with_regions_retained_json`] a step at a time.
    /// This call lowers the body and each [`Self::resume_region_layout`] then
    /// measures up to a number of body blocks, finishing the pass once all are
    /// measured; the layout equals the one-call pass's. Any other region layout,
    /// a document change or a font registration in between abandons it.
    pub fn begin_region_layout(&self, input_json: &str) -> Result<RegionLayoutProgress, String> {
        self.resumable.replace(None);
        let version = self.doc.version();
        let prepared = self.prepare_region_layout(input_json, None)?;
        let progress = self.region_layout_step(ResumableRegionLayout { version, prepared }, 0)?;
        Ok(progress)
    }

    /// Measures up to `blocks` more body blocks of the pass
    /// [`Self::begin_region_layout`] began, and finishes it once all are.
    pub fn resume_region_layout(&self, blocks: usize) -> Result<RegionLayoutProgress, String> {
        let pending = self
            .resumable
            .borrow_mut()
            .take()
            .ok_or_else(|| "no region layout to resume".to_owned())?;
        if pending.version != self.doc.version()
            || pending.prepared.fonts != docx_layout::measure_fonts_generation()
        {
            return Err("the document or its fonts changed since the region layout began".into());
        }
        self.region_layout_step(pending, blocks)
    }

    fn region_layout_step(
        &self,
        mut pending: ResumableRegionLayout,
        blocks: usize,
    ) -> Result<RegionLayoutProgress, String> {
        let body_blocks = pending
            .prepared
            .body
            .as_ref()
            .map_or(pending.prepared.input.measured.len(), |body| {
                body.blocks.len()
            });
        if !pending.prepared.measure(blocks)? {
            let measured_blocks = pending
                .prepared
                .body
                .as_ref()
                .map_or(0, |body| body.flow.measured());
            self.resumable.replace(Some(pending));
            return Ok(RegionLayoutProgress {
                measured_blocks,
                body_blocks,
                layout_json: None,
            });
        }
        let pass = self.finish_region_layout(pending.prepared)?;
        Ok(RegionLayoutProgress {
            measured_blocks: body_blocks,
            body_blocks,
            layout_json: Some(self.retained_region_layout_json(&pass)?),
        })
    }

    fn layout_regions_retained_json(
        &self,
        input_json: &str,
        prefix_pages: Option<usize>,
    ) -> Result<String, String> {
        let pass = self.layout_regions(input_json, prefix_pages)?;
        self.retained_region_layout_json(&pass)
    }

    fn retained_region_layout_json(&self, pass: &RegionPass) -> Result<String, String> {
        let notes_converged = pass.notes_converged;
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
            #[serde(skip_serializing_if = "std::ops::Not::not")]
            provisional: bool,
        }
        serde_json::to_string(&RetainedRegionLayoutOutput {
            layout: pagination
                .layout
                .as_ref()
                .expect("layout retained after successful pagination"),
            headers_footers: state.headers_footers.as_ref(),
            notes_converged,
            provisional: pass.provisional,
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
        Ok(self.layout_regions(input_json, None)?.notes_converged)
    }

    fn layout_regions(
        &self,
        input_json: &str,
        prefix_pages: Option<usize>,
    ) -> Result<RegionPass, String> {
        self.resumable.replace(None);
        let mut prepared = self.prepare_region_layout(input_json, prefix_pages)?;
        prepared.measure(usize::MAX)?;
        self.finish_region_layout(prepared)
    }

    /// A region layout pass through lowering the body; its measurement is left
    /// to [`PreparedRegionLayout::measure`] when the body is measured afresh.
    fn prepare_region_layout(
        &self,
        input_json: &str,
        prefix_pages: Option<usize>,
    ) -> Result<PreparedRegionLayout, String> {
        let request_fingerprint = layout_options_fingerprint(
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?,
        );
        // Reused pages keep the section stamps and page labels of the regions
        // they were laid out under, so a regions change paginates afresh.
        if self
            .regions
            .borrow()
            .as_ref()
            .is_none_or(|state| state.request_fingerprint != request_fingerprint)
        {
            self.pagination.borrow_mut().checkpoints.clear();
        }
        let request: RegionLayoutInput =
            serde_json::from_str(input_json).map_err(|error| format!("parse: {error}"))?;
        let cached_page_totals = request.cached_page_totals;
        let (mut input, mut regions, notes, measurement, render_env, body_story) = request.split();
        let request_options = input.options.clone();
        let mut parsed_render_env = if render_env.is_null() {
            None
        } else {
            Some(
                serde_json::from_value::<RenderEnv>(render_env)
                    .map_err(|error| format!("parse render environment: {error}"))?,
            )
        };
        let revision_preview_key = parsed_render_env
            .as_ref()
            .map_or(Ok(0), revision_preview_key)?;
        if regions.sections.len() <= 1
            && let Some(env) = &mut parsed_render_env
        {
            let line_px = regions.paragraph_spacing_line_px(0);
            env.paragraph_spacing_line_px = (line_px != 16.0).then_some(line_px);
            env.doc_grid_pitch_px = regions.doc_grid_snap_pitch_px(0);
        }
        let fonts = docx_layout::measure_fonts_generation();
        let measurement_fingerprint = serde_json::to_vec(&(
            &measurement.defaults,
            &measurement.compat,
            measurement.authoritative_shaping,
            docx_layout::measure_font_cache_identity(&measurement.font_chains),
        ))
        .map(|bytes| hash_bytes(&bytes))
        .map_err(|error| format!("fingerprint measurement config: {error}"))?;
        let resident_body = body_story.is_some();
        let main_body = body_story.as_deref() == Some("body");
        let mut block_fingerprints: Option<Vec<Fingerprint>> = None;
        let mut lowered_from = None;
        let mut input_lowering = None;
        let mut has_floats = false;
        let mut measured_widths = Vec::new();
        let mut measured_font_dependencies = Vec::new();
        let mut measured_table_wrap_frames = Vec::new();
        let mut measured_float_geometry = None;
        let mut provisional = false;
        let mut body = None;
        if let Some(story) = body_story.as_deref() {
            let render_env = parsed_render_env
                .as_ref()
                .ok_or_else(|| "resident body layout requires a render environment".to_owned())?;
            enum Arena {
                Reused(Vec<MeasuredBlock>, Vec<Fingerprint>),
                /// Blocks to measure, and whether floats couple the whole flow.
                Full(Vec<LayoutBlock>, bool),
            }
            let arena = self
                .with_lowered_story_mapped(
                    story,
                    render_env,
                    &mut || {},
                    |blocks, lowering| -> Result<Arena, String> {
                        input_lowering = Some(lowering);
                        apply_section_geometry(&mut input, &regions);
                        let (widths, table_wrap_frames) =
                            region_measurement_frames(blocks.iter(), &input, &regions);
                        let geometry = initial_float_page_geometry(&input, &regions);
                        measured_widths.clone_from(&widths);
                        measured_table_wrap_frames.clone_from(&table_wrap_frames);
                        measured_float_geometry = Some(float_geometry_key(&geometry));
                        let default_width = widths.first().copied().unwrap_or(0.0);
                        let (floats, margin_floats) =
                            docx_layout::measure_blocks::floating_zone_kinds(
                                blocks,
                                default_width,
                                &measurement,
                                Some(&geometry),
                            )?;
                        has_floats = floats;
                        // margin-relative zones couple the whole flow; re-measure all
                        if margin_floats || (floats && has_wrap_stabilized_shapes(blocks)) {
                            return Ok(Arena::Full(blocks.to_vec(), true));
                        }
                        match self.resident_region_measured(
                            blocks,
                            &widths,
                            &table_wrap_frames,
                            &regions,
                            &measurement,
                            measurement_fingerprint,
                            floats.then_some(&geometry),
                        )? {
                            Some((measured, fingerprints)) => {
                                Ok(Arena::Reused(measured, fingerprints))
                            }
                            None => Ok(Arena::Full(blocks.to_vec(), false)),
                        }
                    },
                )
                .map_err(|error| error.to_string())??;
            if has_floats {
                lowered_from = self
                    .render
                    .borrow()
                    .stories
                    .get(story)
                    .map(|lowered| Rc::clone(&lowered.blocks));
            }
            match arena {
                Arena::Reused(measured, fingerprints) => {
                    measured_font_dependencies =
                        self.pagination.borrow().measured_font_dependencies.clone();
                    input.measured = measured;
                    block_fingerprints = Some(fingerprints);
                    apply_section_geometry(&mut input, &regions);
                }
                Arena::Full(mut blocks, coupled) => {
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
                    let (widths, table_wrap_frames) =
                        region_measurement_frames(blocks.iter(), &input, &regions);
                    let geometry = initial_float_page_geometry(&input, &regions);
                    match prefix_pages {
                        // Floats whose zones only settle later, such as shapes that page-side
                        // wrapping brings into the body, rule a prefix out too.
                        Some(pages) if !coupled && floats_follow_the_text(&blocks) => {
                            let anchored = anchors_objects(&blocks);
                            let (measures, dependencies) = measure_page_prefix(
                                &mut blocks,
                                &widths,
                                &table_wrap_frames,
                                &measurement,
                                &geometry,
                                &request_options,
                                &regions,
                                pages,
                                anchored,
                            )?;
                            measured_font_dependencies = dependencies;
                            provisional = measures.len() < blocks.len();
                            blocks.truncate(measures.len());
                            if provisional {
                                let last = section_breaks(&blocks);
                                input.options =
                                    options_through_section(&request_options, &regions, last);
                                regions.sections.truncate(last + 1);
                            }
                            input.measured = blocks
                                .into_iter()
                                .zip(measures)
                                .map(|(block, measure)| MeasuredBlock { block, measure })
                                .collect();
                        }
                        _ => {
                            let flow =
                                docx_layout::measure_blocks::FloatFlow::with_table_wrap_frames(
                                    &blocks,
                                    &widths,
                                    &table_wrap_frames,
                                    &measurement,
                                    Some(&geometry),
                                )?;
                            body = Some(BodyMeasure {
                                blocks,
                                widths,
                                flow,
                                fingerprints: Vec::new(),
                            });
                        }
                    }
                }
            }
        } else {
            apply_section_geometry(&mut input, &regions);
        }
        Ok(PreparedRegionLayout {
            input_json: input_json.to_owned(),
            request_fingerprint,
            input,
            regions,
            notes,
            measurement,
            parsed_render_env,
            revision_preview_key,
            measurement_fingerprint,
            fonts,
            resident_body,
            main_body,
            block_fingerprints,
            lowered_from,
            input_lowering,
            has_floats,
            measured_widths,
            measured_font_dependencies,
            measured_table_wrap_frames,
            measured_float_geometry,
            provisional,
            cached_page_totals,
            body,
        })
    }

    /// Everything a region layout pass does after body measurement: header and
    /// footer measurement, pagination and notes, and the retained state.
    fn finish_region_layout(&self, prepared: PreparedRegionLayout) -> Result<RegionPass, String> {
        let PreparedRegionLayout {
            input_json,
            request_fingerprint,
            mut input,
            regions,
            mut notes,
            measurement,
            parsed_render_env,
            revision_preview_key,
            measurement_fingerprint,
            fonts,
            resident_body,
            main_body,
            mut block_fingerprints,
            lowered_from,
            input_lowering,
            has_floats,
            measured_widths,
            measured_font_dependencies,
            measured_table_wrap_frames,
            measured_float_geometry,
            provisional,
            cached_page_totals,
            body,
        } = prepared;
        debug_assert!(
            body.is_none(),
            "the body is measured before a pass finishes"
        );
        let body_section_breaks = input
            .measured
            .iter()
            .filter(|measured| matches!(measured.block, LayoutBlock::SectionBreak(_)))
            .count();
        let revision_preview = parsed_render_env
            .as_ref()
            .map(|env| env.revision_preview.clone())
            .unwrap_or_default();
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
            // Header and footer extents widen the section breaks' margins after the
            // reuse walk: fingerprint them as a full pass does, as paginated.
            Some(mut fingerprints) => {
                for (fingerprint, measured) in fingerprints.iter_mut().zip(&input.measured) {
                    if matches!(measured.block, LayoutBlock::SectionBreak(_)) {
                        *fingerprint = measured_fingerprint(measured)?;
                    }
                }
                fingerprints
            }
            None => measured_fingerprints(&input)?,
        };
        let refs = input
            .measured
            .iter()
            .flat_map(|measured| collect_note_refs(std::slice::from_ref(&measured.block)))
            .collect::<Vec<_>>();
        let previous_notes = note_page_keys(self.pagination.borrow().layout.as_ref());
        // The note fixpoint replays `base_input`. Without notes `input` is the
        // final pass; with notes the final pass carries reserved heights, so
        // the reservation-free pass stays out of the retained pagination state
        // that the next edit paginates against.
        // Placement only zeroes contextual spacing, which every pass applies
        // again, so the note passes replay the body arena in place. Page-side
        // wrapping rewrites shapes per pass, so it replays a copy instead.
        let mut separator_heights = input
            .options
            .note_separator_heights
            .clone()
            .unwrap_or_default();
        let mut base_input = (!refs.is_empty()
            && resident_body
            && input.measured.iter().any(|measured| {
                matches!(&measured.block, LayoutBlock::Shape(shape) if wraps_by_page_side(shape))
            }))
        .then(|| input.clone());
        let mut restamped_pages = Vec::new();
        let (mut initial_layout, mut arena) = if refs.is_empty() {
            self.layout_document_value_with_fingerprints(
                input,
                block_fingerprints.clone(),
                Some((revision_preview_key, &revision_preview, main_body)),
                cached_page_totals,
            )?;
            let layout = self
                .pagination
                .borrow_mut()
                .layout
                .take()
                .expect("layout retained after successful pagination");
            (layout, None)
        } else {
            let layout =
                docx_layout::place::layout_document(&mut input).map_err(layout_error_message)?;
            (layout, Some(input))
        };
        apply_document_regions_tracked(&mut initial_layout, &regions, &mut restamped_pages);
        let presentations = build_note_presentations(&refs, &initial_layout.pages, &regions);
        assign_note_presentations(&mut notes.contents, &presentations);
        if resident_body {
            let measured_separator_heights = self.measure_resident_notes(
                &mut notes.contents,
                &refs,
                &initial_layout,
                &regions,
                &measurement,
                parsed_render_env
                    .as_ref()
                    .expect("resident body required render environment"),
            )?;
            separator_heights
                .footnote
                .extend(measured_separator_heights.footnote);
            separator_heights
                .endnote
                .extend(measured_separator_heights.endnote);
        }
        if separator_heights != NoteSeparatorHeights::default() {
            for input in base_input.iter_mut().chain(arena.iter_mut()) {
                input.options.note_separator_heights = Some(separator_heights.clone());
            }
        }
        let stabilized = stabilize_note_layout(
            |reserved| {
                let mut copy;
                let pass = match (&base_input, arena.as_mut()) {
                    (Some(base), _) => {
                        copy = base.clone();
                        &mut copy
                    }
                    (None, Some(arena)) => arena,
                    (None, None) => {
                        return Err(docx_layout::LayoutError::Invalid(
                            "note passes without a body arena".to_owned(),
                        ));
                    }
                };
                pass.options.footnote_reserved_heights = reservation_options(reserved);
                if resident_body {
                    stabilize_shape_wrapping(pass, &regions, &measurement)?;
                }
                let mut layout = docx_layout::place::layout_document(pass)?;
                apply_document_regions_tracked(&mut layout, &regions, &mut restamped_pages);
                Ok(layout)
            },
            &refs,
            &notes.contents,
            &separator_heights,
            initial_layout,
            &regions,
        )
        .map_err(layout_error_message)?;
        let notes_converged = stabilized.converged;
        if let Some(mut final_input) = base_input.or(arena) {
            final_input.options.footnote_reserved_heights =
                reservation_options(&stabilized.reserved_heights);
            let reshaped = resident_body
                && stabilize_shape_wrapping(&mut final_input, &regions, &measurement)
                    .map_err(layout_error_message)?;
            let fingerprints = if reshaped {
                measured_fingerprints(&final_input)?
            } else {
                block_fingerprints
            };
            self.layout_document_value_with_fingerprints(
                final_input,
                fingerprints,
                Some((revision_preview_key, &revision_preview, main_body)),
                cached_page_totals,
            )?;
            restamped_pages.clear();
        } else {
            self.pagination.borrow_mut().layout = Some(stabilized.layout);
        }
        let mut pagination = self.pagination.borrow_mut();
        let layout = pagination
            .layout
            .as_mut()
            .expect("layout retained after successful pagination");
        // Pages an incremental pass reused still carry the previous pass's notes.
        for page in &mut layout.pages {
            page.footnote_ids = None;
            page.footnote_columns = None;
            page.note_areas = None;
        }
        layout.partial = provisional || self.partial_document.get();
        layout.cached_page_totals = cached_page_totals;
        apply_document_regions_tracked(layout, &regions, &mut restamped_pages);
        let page_note_map = map_notes_to_pages(&layout.pages, &refs, &regions);
        stamp_note_pages(layout, &page_note_map, &regions);
        attach_note_areas(
            layout,
            &page_note_map,
            &notes.contents,
            &separator_heights,
            &regions,
        );
        let note_settlement = if notes_converged {
            NoteSettlement::Converged
        } else {
            let required = calculate_note_reserved_heights(
                &page_note_map,
                &notes.contents,
                &layout.pages,
                &separator_heights,
                &regions,
            );
            match reservation_surplus_pages(&stabilized.reserved_heights, &required) {
                Some(numbers) => NoteSettlement::Covering(
                    layout
                        .pages
                        .iter()
                        .enumerate()
                        .filter(|(_, page)| numbers.contains(&page.number))
                        .map(|(index, _)| index)
                        .collect(),
                ),
                None => NoteSettlement::Unsettled,
            }
        };
        let note_changed_pages: Vec<usize> = note_page_keys(Some(layout))
            .iter()
            .enumerate()
            .filter(|(index, keys)| previous_notes.get(*index) != Some(*keys))
            .map(|(index, _)| index)
            .collect();
        let measured_value = measured_headers_footers
            .as_mut()
            .map(|payload| {
                resolve_header_footer_field_widths(payload, layout, &measurement)?;
                serde_json::to_value(&*payload)
                    .map_err(|error| format!("serialize headers/footers: {error}"))
            })
            .transpose()?;
        pagination.note_changed_pages = note_changed_pages;
        if let Some(pages) = &mut pagination.restamped_pages {
            pages.extend(restamped_pages);
        }
        let serial = pagination.layout_epoch;
        let headers_footers = measured_value.or_else(|| regions.headers_footers.clone());
        let notes_clear = notes.contents.is_empty() && refs.is_empty();
        // Multi-section documents are excluded: an edit can move a section
        // boundary without changing the total page count, which changes
        // section-relative page labels and the PAGE/NUMPAGES field widths
        // baked into the retained headers/footers payload. With one section,
        // an unchanged page count implies unchanged labels.
        let single_section = self.lays_out_as_one_section(&regions, body_section_breaks);
        drop(pagination);
        let regional = match (
            resident_body && single_section && !provisional,
            parsed_render_env.as_ref(),
        ) {
            (true, Some(env)) => Some((self.regional_fingerprint(&regions, env), env.clone())),
            _ => None,
        };
        self.regions.replace(Some(ResidentRegionState {
            request_json: input_json,
            request_fingerprint,
            headers_footers,
            notes_converged,
            provisional,
            fast_path: regional.map(|(regional, render_env)| RegionFastPathState {
                cached_page_totals,
                regions: Rc::new(regions),
                measurement: Rc::new(measurement),
                measurement_fingerprint,
                fonts,
                regional,
                notes_clear,
                render_env,
            }),
        }));
        // Only the region-measured arena may seed the next pass's reuse walk.
        let mut pagination = self.pagination.borrow_mut();
        pagination.measured_with =
            (resident_body && !provisional).then_some(measurement_fingerprint);
        pagination.lowered_from = lowered_from;
        pagination.input_lowering = input_lowering;
        pagination.measured_widths = measured_widths;
        pagination.measured_font_dependencies = measured_font_dependencies;
        pagination.measured_table_wrap_frames = measured_table_wrap_frames;
        pagination.measured_float_geometry = measured_float_geometry;
        pagination.measured_with_floats = has_floats;
        // Checkpoints of a prefix pass describe a cut document.
        if provisional {
            pagination.checkpoints.clear();
        }
        drop(pagination);
        if let (true, Some(render_env)) = (resident_body && !provisional, parsed_render_env) {
            self.capture.replace(Some(LayoutCapture {
                version: self.doc.version(),
                serial,
                fonts,
                note_settlement,
                render_env,
                headers_footers: measured_headers_footers.map(Rc::new),
                notes: Rc::new(notes.contents),
            }));
        }
        Ok(RegionPass {
            notes_converged,
            provisional,
        })
    }

    /// Whether a body with `section_breaks` lays out as one section of `regions`. Hosts repeat
    /// the final section after the parsed ones, which local lowering reads as one section when
    /// the body has no section break.
    fn lays_out_as_one_section(&self, regions: &DocumentRegions, section_breaks: usize) -> bool {
        regions.sections.len() <= 1 || (self.local_lowering.get() && section_breaks == 0)
    }

    /// A fingerprint of the header and footer stories `regions` reference, lowered in `env`,
    /// cell and control content included; a story that cannot be lowered counts as absent.
    fn regional_fingerprint(&self, regions: &DocumentRegions, env: &RenderEnv) -> u64 {
        let mut stories = BTreeSet::new();
        for section_index in 0..regions.sections.len() {
            if let Some(refs) = effective_header_footer_refs(regions, section_index) {
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
        }
        let mut bytes = Vec::new();
        for story in stories {
            let lowered = self
                .with_lowered_story(&story, env, |blocks| serde_json::to_vec(blocks).ok())
                .ok()
                .flatten();
            bytes.extend_from_slice(story.as_bytes());
            bytes.push(0);
            bytes.extend_from_slice(
                &hash_bytes(lowered.as_deref().unwrap_or_default()).to_le_bytes(),
            );
            bytes.push(u8::from(lowered.is_some()));
        }
        hash_bytes(&bytes)
    }

    fn lower_note_separator(
        &self,
        kind: &str,
        render_env: &RenderEnv,
    ) -> Result<Option<Rc<Vec<LayoutBlock>>>, String> {
        let Some(state) = self.doc.note_separator_state()? else {
            self.note_separators.borrow_mut().take();
            return Ok(None);
        };
        let mut cache = self.note_separators.borrow_mut();
        if cache
            .as_ref()
            .is_none_or(|cached| !Arc::ptr_eq(&cached.state, &state) || cached.env != *render_env)
        {
            let scratch = EditingDoc::new(1);
            scratch
                .apply_update_v1(&state)
                .map_err(|error| error.to_string())?;
            *cache = Some(LoweredNoteSeparators {
                state,
                doc: scratch,
                env: render_env.clone(),
                blocks: HashMap::new(),
                revealable: HashMap::new(),
            });
        }
        let cached = cache.as_mut().expect("separator cache initialized");
        if let Some(blocks) = cached.blocks.get(kind) {
            return Ok(Some(Rc::clone(blocks)));
        }
        if cached.doc.story_len(kind).is_err() {
            return Ok(None);
        }
        let blocks = Rc::new(
            crate::bridge::yrs_doc_to_layout_blocks(&cached.doc, kind, render_env)
                .map_err(|error| error.to_string())?,
        );
        cached.blocks.insert(kind.to_owned(), Rc::clone(&blocks));
        Ok(Some(blocks))
    }

    /// The separator blocks a revision preview can reveal, for the preview font superset.
    fn note_separator_revealable(
        &self,
        kind: &str,
        render_env: &RenderEnv,
    ) -> Result<Option<Rc<Vec<LayoutBlock>>>, String> {
        if self.lower_note_separator(kind, render_env)?.is_none() {
            return Ok(None);
        }
        let mut cache = self.note_separators.borrow_mut();
        let cached = cache.as_mut().expect("separator cache initialized");
        if let Some(blocks) = cached.revealable.get(kind) {
            return Ok(Some(Rc::clone(blocks)));
        }
        let mut local = crate::bridge::local::LocalLowering::new(false);
        let (_, _, revealable, _) = crate::bridge::preview::lower_recorded(
            &cached.doc,
            kind,
            render_env,
            &mut local,
            false,
        )
        .map_err(|error| error.to_string())?;
        let revealable = Rc::new(revealable);
        cached
            .revealable
            .insert(kind.to_owned(), Rc::clone(&revealable));
        Ok(Some(revealable))
    }

    fn measure_note_separators(
        &self,
        contents: &[docx_layout::footnotes::NoteContent],
        layout: &Layout,
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        render_env: &RenderEnv,
    ) -> Result<NoteSeparatorHeights, String> {
        use docx_layout::paragraph_spacing::{
            apply_contextual_spacing_blocks, get_spacing_after, get_spacing_before,
        };

        let mut heights = NoteSeparatorHeights::default();
        let Some(first_page) = layout.pages.first() else {
            return Ok(heights);
        };
        let sections: BTreeSet<_> = (0..regions.sections.len())
            .chain(layout.pages.iter().map(|page| page.region_section_index))
            .collect();
        let mut first_pages = HashMap::new();
        for page in &layout.pages {
            first_pages.entry(page.region_section_index).or_insert(page);
        }
        for (kind, name) in [
            (NoteKind::Footnote, "footnote"),
            (NoteKind::Endnote, "endnote"),
        ] {
            if !contents.iter().any(|content| content.note_kind == kind) {
                continue;
            }
            let Some(lowered) = self.lower_note_separator(name, render_env)? else {
                continue;
            };
            let mut measured_heights = HashMap::new();
            for &section_index in &sections {
                let page = first_pages.get(&section_index).copied();
                let section = regions
                    .sections
                    .get(section_index)
                    .or_else(|| regions.sections.last());
                let size = page
                    .map(|page| &page.size)
                    .or_else(|| section.and_then(|section| section.page_size.as_ref()))
                    .unwrap_or(&first_page.size);
                let margins = page
                    .map(|page| page.body_margins.as_ref().unwrap_or(&page.margins))
                    .or_else(|| section.and_then(|section| section.margins.as_ref()))
                    .unwrap_or(&first_page.margins);
                let width = (size.w - margins.left - margins.right).max(1.0);
                let line_px = regions.paragraph_spacing_line_px(section_index);
                let grid_pitch = regions.doc_grid_snap_pitch_px(section_index);
                let key = (
                    width.to_bits(),
                    line_px.to_bits(),
                    grid_pitch.map(f64::to_bits),
                );
                let height = if let Some(&height) = measured_heights.get(&key) {
                    height
                } else {
                    let mut blocks = lowered.as_ref().clone();
                    for block in &mut blocks {
                        resolve_line_unit_spacing(block, line_px);
                        resolve_doc_grid_pitch(block, grid_pitch);
                    }
                    apply_contextual_spacing_blocks(&mut blocks);
                    let measures = docx_layout::measure_blocks::measure_blocks(
                        &mut blocks,
                        width,
                        measurement,
                    )?;
                    let mut height = 0.0;
                    let mut previous_after = 0.0_f64;
                    for (block, measure) in blocks.iter().zip(&measures) {
                        if let (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent)) =
                            (block, measure)
                        {
                            height += previous_after.max(get_spacing_before(paragraph))
                                + extent
                                    .lines
                                    .iter()
                                    .map(|line| line.line_height)
                                    .sum::<f64>();
                            previous_after = get_spacing_after(paragraph);
                        } else {
                            height += previous_after
                                + docx_layout::measure_blocks::extent_height(measure);
                            previous_after = 0.0;
                        }
                    }
                    height += previous_after;
                    measured_heights.insert(key, height);
                    height
                };
                match kind {
                    NoteKind::Footnote => &mut heights.footnote,
                    NoteKind::Endnote => &mut heights.endnote,
                }
                .insert(section_index, height);
            }
        }
        Ok(heights)
    }

    fn measure_resident_notes(
        &self,
        contents: &mut [docx_layout::footnotes::NoteContent],
        refs: &[docx_layout::footnotes::NoteRefLocation],
        layout: &Layout,
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        render_env: &RenderEnv,
    ) -> Result<NoteSeparatorHeights, String> {
        let separator_heights =
            self.measure_note_separators(contents, layout, regions, measurement, render_env)?;
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
        Ok(separator_heights)
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
        self.layout_document_value_with_fingerprints(input, block_fingerprints, None, false)
    }

    /// Paginate a resident measured arena whose clean block fingerprints were
    /// retained while rebuilding the dirty paragraph. Compatibility callers
    /// still enter through `layout_document_value` and fingerprint every block.
    /// The paragraphs of every story that hold a revision whose previewed
    /// decision differs between `previous` and `next`. Story ids are opaque, so
    /// a body table's cells may live under any id.
    fn preview_changed_paragraphs(
        &self,
        previous: &BTreeMap<String, RevisionPreview>,
        next: &BTreeMap<String, RevisionPreview>,
    ) -> Result<HashSet<String>, String> {
        let changed: HashSet<&str> = previous
            .keys()
            .chain(next.keys())
            .filter(|id| previous.get(*id) != next.get(*id))
            .map(String::as_str)
            .collect();
        let mut paragraphs = HashSet::new();
        if changed.is_empty() {
            return Ok(paragraphs);
        }
        use yrs::{Map, ReadTxn, Transact};
        let stories: Vec<String> = {
            let txn = self.doc.yrs_doc().transact();
            txn.get_map(crate::STORIES)
                .map_or_else(Vec::new, |stories| {
                    stories
                        .iter(&txn)
                        .filter(|(_, value)| value.clone().cast::<yrs::TextRef>().is_ok())
                        .map(|(story, _)| story.to_owned())
                        .collect()
                })
        };
        for story in stories {
            let txn = self.doc.yrs_doc().transact();
            let text = crate::story_ref(&txn, &story).map_err(|error| error.to_string())?;
            let mut ranges: Vec<(u32, u32)> = self
                .doc
                .story_raw_changes(&story, &text, &txn)
                .into_iter()
                .filter(|change| changed.contains(change.id.as_str()))
                .map(|change| (change.start, change.end))
                .collect();
            // An inline content control keeps its content's revisions in its payload.
            ranges.extend(
                self.doc
                    .chunk_snapshot(&story, &text, &txn)
                    .iter()
                    .filter(|chunk| {
                        matches!(&chunk.kind, crate::ops::ChunkKind::Embed(Some(map))
                            if crate::map_string(map, &txn, crate::KIND_KEY).as_deref() == Some("sdt"))
                    })
                    .map(|chunk| (chunk.start, chunk.end())),
            );
            if ranges.is_empty() {
                continue;
            }
            // Paragraphs come in story order: one overlaps a range when the
            // furthest end among the ranges starting by its pilcrow reaches it.
            ranges.sort_unstable();
            let mut next_range = 0;
            let mut furthest_end = None;
            for bounds in crate::op::para_bounds(&text, &txn) {
                while let Some(&(from, to)) = ranges.get(next_range)
                    && from <= bounds.pilcrow
                {
                    furthest_end = furthest_end.max(Some(to));
                    next_range += 1;
                }
                if furthest_end.is_some_and(|end| bounds.start <= end) {
                    paragraphs.insert(bounds.para_id);
                }
            }
        }
        Ok(paragraphs)
    }

    fn layout_document_value_with_fingerprints(
        &self,
        mut input: LayoutInput,
        mut block_fingerprints: Vec<Fingerprint>,
        revision_preview: Option<(u64, &BTreeMap<String, RevisionPreview>, bool)>,
        cached_page_totals: bool,
    ) -> Result<(), String> {
        if block_fingerprints.len() != input.measured.len() {
            return Err("resident pagination fingerprints do not match measured blocks".to_owned());
        }
        self.resumable.replace(None);
        self.capture.borrow_mut().take();
        let input_options_fingerprint = options_fingerprint(&input)?;
        let previous_page_count = self
            .pagination
            .borrow()
            .layout
            .as_ref()
            .map(|layout| layout.pages.len());
        let mut incremental = false;
        let mut deltas = HashMap::new();
        let run = {
            let mut previous = self.pagination.borrow_mut();
            if let Some(previous_input) = previous.input.as_ref() {
                for ((previous, next), fingerprint) in previous_input
                    .measured
                    .iter()
                    .zip(&input.measured)
                    .zip(&mut block_fingerprints)
                {
                    if relative_run_position_fingerprint(&previous.block)
                        != relative_run_position_fingerprint(&next.block)
                    {
                        *fingerprint = measured_fingerprint(next)?;
                    }
                }
            }
            // A preview shows the document at its source positions, so a block whose
            // content is unchanged but whose positions moved shows other source: it is
            // placed afresh rather than shifted.
            // Over an unchanged document lowered from the body story, only a block
            // holding a revision whose decision changed can show other source.
            if let Some((key, preview, resident)) = revision_preview
                && key != previous.revision_preview_key
                && incremental_eligible(&previous, &input, input_options_fingerprint)
                && let moved = (resident && previous.doc_epoch == self.doc_epoch())
                    .then(|| self.preview_changed_paragraphs(&previous.revision_preview, preview))
                    .transpose()?
                && moved.as_ref().is_none_or(|moved| !moved.is_empty())
            {
                let retained = previous.input.as_ref().expect("eligibility checked input");
                for (((fingerprint, retained_fingerprint), next), retained) in block_fingerprints
                    .iter_mut()
                    .zip(&previous.block_fingerprints)
                    .zip(&input.measured)
                    .zip(&retained.measured)
                {
                    if *fingerprint == *retained_fingerprint
                        && moved
                            .as_ref()
                            .is_none_or(|moved| block_holds_paragraph(&next.block, moved))
                        && crate::fingerprint::fingerprint_with_positions(&next.block)?
                            != crate::fingerprint::fingerprint_with_positions(&retained.block)?
                    {
                        *fingerprint ^= 1;
                    }
                }
            }
            let first_dirty = previous
                .block_fingerprints
                .iter()
                .zip(&block_fingerprints)
                .position(|(previous, next)| previous != next)
                .map(|dirty| {
                    section_start_of_first_changed_break(
                        &input.measured,
                        &previous.block_fingerprints,
                        &block_fingerprints,
                        dirty,
                    )
                });
            if let Some(dirty_index) = first_dirty
                && incremental_eligible(&previous, &input, input_options_fingerprint)
            {
                let previous = &mut *previous;
                deltas = position_deltas(
                    previous.input.as_ref().expect("eligibility checked input"),
                    &input,
                );
                let attempted = docx_layout::place::layout_document_incremental_ranges(
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
                    Err(docx_layout::LayoutError::Unsupported(_)) => full_pass(&mut input)?,
                    Err(docx_layout::LayoutError::Invalid(reason)) => return Err(reason),
                }
            } else {
                full_pass(&mut input)?
            }
        };
        let docx_layout::place::IncrementalLayout {
            checkpointed: run,
            rebuilt_page_ranges,
        } = run;
        let mut pagination = self.pagination.borrow_mut();
        pagination.input = Some(input);
        pagination.measured_with = None;
        pagination.measured_font_dependencies.clear();
        pagination.lowered_from = None;
        pagination.input_lowering = None;
        pagination.note_changed_pages.clear();
        let mut layout = run.layout;
        // Every pass over part of a package, the resident edit paths' too.
        layout.partial = self.partial_document.get();
        layout.cached_page_totals = cached_page_totals;
        let same_page_count = previous_page_count == Some(layout.pages.len());
        pagination.layout = Some(layout);
        pagination.checkpoints = run.checkpoints;
        pagination.block_fingerprints = block_fingerprints;
        pagination.options_fingerprint = input_options_fingerprint;
        pagination.doc_epoch = self.doc_epoch();
        if let Some((key, preview, _)) = revision_preview {
            pagination.revision_preview_key = key;
            pagination.revision_preview = preview.clone();
        }
        pagination.rebuilt_page_start = run.rebuilt_page_start;
        pagination.rebuilt_page_end = run.rebuilt_page_end;
        pagination.rebuilt_page_ranges = rebuilt_page_ranges;
        pagination.position_deltas = deltas;
        pagination.last_incremental = incremental;
        if !incremental || !same_page_count {
            pagination.restamped_pages = None;
        }
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
    /// template after committing the text. A table cell or other story nested
    /// in the body qualifies once the body lays out through regions.
    pub fn can_apply_input(&self, story: &str, para_id: &str) -> bool {
        let regions = self.regions.borrow().is_some();
        if story != "body" && !(regions && is_nested_body_story(story)) {
            return false;
        }
        let render_ready = self.render.borrow().stories.contains_key("body");
        let pagination = self.pagination.borrow();
        let layout_ready = pagination.input.is_some() && pagination.layout.is_some();
        drop(pagination);
        let display_ready = self.display.borrow().extras_json.is_some();
        let measure_ready = regions
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
        self.with_lowered_story_mapped(story, &env, after_lower, |blocks, lowering| {
            let mut resident = self.resident_layout_input_from_blocks(
                blocks,
                false,
                false,
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
                        serde_json::to_value(&*next_block)
                            .map_err(|error| format!("serialize dirty paragraph: {error}"))?,
                    );
                    let envelope_json = serde_json::to_string(&envelope)
                        .map_err(|error| format!("serialize measurement envelope: {error}"))?;
                    let extent_json = docx_layout::measure_paragraph_json_resident(&envelope_json)?;
                    let extent: ParagraphExtent = serde_json::from_str(&extent_json)
                        .map_err(|error| format!("parse resident paragraph extent: {error}"))?;
                    Ok(BlockExtent::Paragraph(extent))
                },
            )?;
            resident.lowering = Some(lowering);
            Ok(resident)
        })
        .map_err(|error| error.to_string())?
    }

    /// Shared dirty-block walk over a freshly lowered story: structurally clean
    /// blocks reuse their retained extents (with fresh absolute positions),
    /// changed paragraph blocks are re-measured through `measure_dirty`
    /// (`(block_index, block_key, previous_block, next_block) -> extent`).
    /// With `any_block`, a changed table or other non-paragraph block with a
    /// stable id is re-measured too instead of refused.
    fn resident_layout_input_from_blocks(
        &self,
        blocks: &[LayoutBlock],
        any_block: bool,
        take: bool,
        measure_dirty: &mut dyn FnMut(
            usize,
            &str,
            &LayoutBlock,
            &mut LayoutBlock,
        ) -> Result<BlockExtent, String>,
    ) -> Result<ResidentLayoutInput, String> {
        let mut pagination = self.pagination.borrow_mut();
        let pagination = &mut *pagination;
        let previous = pagination
            .input
            .as_mut()
            .ok_or_else(|| "resident pagination input is not built".to_owned())?;
        let previous_fingerprints = &pagination.block_fingerprints;
        let paragraph_merge = blocks.len().checked_add(1) == Some(previous.measured.len());
        if blocks.len() != previous.measured.len() && !paragraph_merge {
            return Err("resident plain-text input changed the block structure".to_owned());
        }
        if previous.measured.len() != previous_fingerprints.len() {
            return Err("resident pagination fingerprints are not built".to_owned());
        }

        let mut measured = Vec::with_capacity(blocks.len());
        let mut block_fingerprints = Vec::with_capacity(blocks.len());
        let mut moved = Vec::new();
        let mut font_dependencies = Vec::with_capacity(blocks.len());
        let walked = resident_walk(
            blocks,
            any_block,
            paragraph_merge,
            (
                &mut previous.measured,
                previous_fingerprints,
                &pagination.measured_font_dependencies,
                take,
            ),
            measure_dirty,
            (
                &mut measured,
                &mut block_fingerprints,
                &mut moved,
                &mut font_dependencies,
            ),
        );
        let (resident_measure_calls, resident_reused_blocks) = match walked {
            Ok(counts) => counts,
            Err(error) => {
                for (from, to) in moved {
                    previous.measured[from].measure =
                        std::mem::replace(&mut measured[to].measure, BlockExtent::Unsupported);
                }
                return Err(error);
            }
        };

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
            font_dependencies,
            lowering: None,
        })
    }

    /// Reuse retained extents for blocks that cannot have changed (equal
    /// normalized block, width, config, and section-break adjacency).
    /// `Ok(None)` means the caller must measure the whole story.
    /// With `floats`, a changed block re-measures every block of its float
    /// flow segment instead of itself alone.
    #[allow(clippy::too_many_arguments)]
    fn resident_region_measured(
        &self,
        blocks: &[LayoutBlock],
        widths: &[f64],
        table_wrap_frames: &[bool],
        regions: &DocumentRegions,
        measurement: &docx_layout::measure_blocks::MeasurementConfig,
        measurement_fingerprint: u64,
        floats: Option<&docx_layout::measure_blocks::FloatPageGeometry>,
    ) -> Result<Option<(Vec<MeasuredBlock>, Vec<Fingerprint>)>, String> {
        let pagination = &mut *self.pagination.borrow_mut();
        let Some(previous) = pagination.input.as_mut() else {
            return Ok(None);
        };
        if pagination.measured_with != Some(measurement_fingerprint)
            || previous.measured.len() != blocks.len()
            || previous.measured.len() != pagination.block_fingerprints.len()
            || previous.measured.len() != pagination.measured_font_dependencies.len()
        {
            return Ok(None);
        }
        let previous_widths = &pagination.measured_widths;
        if previous_widths.len() != blocks.len()
            || (pagination.measured_with_floats && floats.is_none())
        {
            return Ok(None);
        }
        let lowered_from = match floats {
            Some(_) => match pagination.lowered_from.as_deref() {
                Some(lowered) if lowered.len() == blocks.len() => Some(lowered),
                _ => return Ok(None),
            },
            None => None,
        };
        if floats.is_some_and(|geometry| {
            widths.first() != previous_widths.first()
                || table_wrap_frames != pagination.measured_table_wrap_frames.as_slice()
                || pagination.measured_float_geometry != Some(float_geometry_key(geometry))
        }) {
            return Ok(None);
        }
        let previous_fingerprints = &pagination.block_fingerprints;
        let mut dependencies = pagination.measured_font_dependencies.clone();
        let float_blocks = if floats.is_some() { blocks.len() } else { 0 };
        let mut float_dirty = vec![false; float_blocks];
        let mut float_sections = vec![0; float_blocks];

        let mut measured: Vec<MeasuredBlock> = Vec::with_capacity(blocks.len());
        let mut block_fingerprints = Vec::with_capacity(blocks.len());
        let mut measure_calls = 0_u64;
        let mut reused_blocks = 0_u64;
        let mut section_index = 0_usize;
        for (index, next_block) in blocks.iter().enumerate() {
            let section_break_mark = docx_layout::measure_blocks::is_section_break_mark(
                next_block,
                index
                    .checked_sub(1)
                    .and_then(|previous| blocks.get(previous)),
                blocks.get(index + 1),
            );
            let retained_section_break_mark = docx_layout::measure_blocks::is_section_break_mark(
                &previous.measured[index].block,
                index
                    .checked_sub(1)
                    .and_then(|index| previous.measured.get(index))
                    .map(|measured| &measured.block),
                previous
                    .measured
                    .get(index + 1)
                    .map(|measured| &measured.block),
            );
            if let Some(section) = float_sections.get_mut(index) {
                *section = section_index;
            }
            let previous_entry = &mut previous.measured[index];
            if !resident_block_slots_match(&previous_entry.block, next_block) {
                restore_moved_measures(&mut previous.measured, measured);
                return Ok(None);
            }
            // In a float flow the extent depends on the block as lowered, before
            // contextual spacing, so that form has to be unchanged too.
            let width_clean = dependencies[index]
                .matches(FontChains::BTree(&measurement.font_chains))
                && section_break_mark == retained_section_break_mark
                && widths.get(index) == previous_widths.get(index)
                && matches!(previous_entry.measure, BlockExtent::Unsupported)
                    == matches!(next_block, LayoutBlock::Unsupported)
                && lowered_from.is_none_or(|lowered| lowered[index] == *next_block);
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
                if width_clean
                    && (owned == previous_entry.block
                        || section_breaks_match_but_margins(&owned, &previous_entry.block))
                {
                    measured.push(MeasuredBlock {
                        block: owned,
                        measure: std::mem::replace(
                            &mut previous_entry.measure,
                            BlockExtent::Unsupported,
                        ),
                    });
                    block_fingerprints.push(previous_fingerprints[index]);
                    reused_blocks = reused_blocks.wrapping_add(1);
                } else if floats.is_some() {
                    float_dirty[index] = true;
                    measured.push(MeasuredBlock {
                        block: owned,
                        measure: previous_entry.measure.clone(),
                    });
                    block_fingerprints.push(previous_fingerprints[index]);
                } else {
                    let (measure, reads) = FontChainDependencies::capture(|| {
                        if section_break_mark {
                            Ok(BlockExtent::Paragraph(ParagraphExtent {
                                lines: Vec::new(),
                                total_height: 0.0,
                            }))
                        } else {
                            docx_layout::measure_blocks::measure_block(
                                &mut owned,
                                widths.get(index).copied().unwrap_or(0.0),
                                measurement,
                            )
                        }
                    });
                    let measure = match measure {
                        Ok(measure) => measure,
                        Err(error) => {
                            restore_moved_measures(&mut previous.measured, measured);
                            return Err(error);
                        }
                    };
                    dependencies[index] = reads;
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
        if let Some(geometry) = floats {
            let marks = docx_layout::measure_blocks::section_break_marks(blocks);
            let default_width = widths.first().copied().unwrap_or(0.0);
            let mut start = 0;
            while start < blocks.len() {
                let end = (start + 1..blocks.len())
                    .find(|&index| docx_layout::measure_blocks::resets_float_flow(&blocks[index]))
                    .unwrap_or(blocks.len());
                if float_dirty[start..end].contains(&true) {
                    // Measured in the full pass's form, before contextual spacing.
                    let mut segment = blocks[start..end].to_vec();
                    for (block, &section) in segment.iter_mut().zip(&float_sections[start..end]) {
                        resolve_line_unit_spacing(
                            block,
                            regions.paragraph_spacing_line_px(section),
                        );
                        resolve_doc_grid_pitch(block, regions.doc_grid_snap_pitch_px(section));
                        if let LayoutBlock::SectionBreak(section_break) = block
                            && let Some(section) = regions.sections.get(section)
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
                    }
                    let extents =
                        docx_layout::measure_blocks::measure_float_segment_with_font_dependencies(
                            &mut segment,
                            &widths[start..end],
                            default_width,
                            &table_wrap_frames[start..end],
                            measurement,
                            Some(geometry),
                            &marks[start..end],
                        );
                    let (extents, reads) = match extents {
                        Ok(Some(measured)) => measured,
                        outcome => {
                            restore_moved_measures(&mut previous.measured, measured);
                            return outcome.map(|_| None);
                        }
                    };
                    for (offset, (mut block, measure)) in
                        segment.into_iter().zip(extents).enumerate()
                    {
                        suppress_contextual_spacing(blocks, start + offset, &mut block);
                        docx_layout::paragraph_spacing::apply_contextual_spacing_blocks(
                            std::slice::from_mut(&mut block),
                        );
                        let entry = MeasuredBlock { block, measure };
                        block_fingerprints[start + offset] = match measured_fingerprint(&entry) {
                            Ok(fingerprint) => fingerprint,
                            Err(error) => {
                                restore_moved_measures(&mut previous.measured, measured);
                                return Err(error);
                            }
                        };
                        measured[start + offset] = entry;
                        dependencies[start + offset] = reads[offset].clone();
                        measure_calls = measure_calls.wrapping_add(1);
                    }
                }
                start = end;
            }
        }
        pagination.measured_font_dependencies = dependencies;
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
        self.layout_document_value_with_fingerprints(
            resident.input,
            resident.block_fingerprints,
            None,
            false,
        )?;
        self.pagination.borrow_mut().input_lowering = resident.lowering;
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
        if story != "body" && !is_nested_body_story(story) {
            return Ok(false);
        }
        let story = "body";
        let fast_config = {
            let state = self.regions.borrow();
            state.as_ref().and_then(|state| {
                let fast = state.fast_path.as_ref()?;
                (fast.notes_clear && fast.fonts == docx_layout::measure_fonts_generation()).then(
                    || {
                        (
                            Rc::clone(&fast.regions),
                            Rc::clone(&fast.measurement),
                            fast.measurement_fingerprint,
                            fast.regional,
                            fast.render_env.clone(),
                            fast.cached_page_totals,
                        )
                    },
                )
            })
        };
        let Some((
            regions,
            measurement,
            measurement_fingerprint,
            regional,
            pass_env,
            cached_page_totals,
        )) = fast_config
        else {
            return Ok(false);
        };
        let env = {
            let render = self.render.borrow();
            let Some(lowered) = render.stories.get(story) else {
                return Ok(false);
            };
            lowered.env.clone()
        };
        if env != pass_env {
            return Ok(false);
        }
        if self.regional_fingerprint(&regions, &env) != regional
            || revision_preview_key(&env)? != self.pagination.borrow().revision_preview_key
        {
            return Ok(false);
        }
        let misses = self.render.borrow().cache_misses;
        let outcome = self
            .with_lowered_story_mapped(
                story,
                &env,
                &mut || phase(RegionResidentPhase::Lowered),
                |blocks,
                 lowering|
                 -> Result<Option<(ResidentLayoutInput, usize, Vec<f64>)>, String> {
                    // A repeated final section takes this path only for a locally patched body.
                    if regions.sections.len() > 1 && self.render.borrow().cache_misses != misses {
                        return Ok(None);
                    }
                    let (widths, geometry, previous_pages) = {
                        let pagination = self.pagination.borrow();
                        let (Some(input), Some(layout), false) = (
                            pagination.input.as_ref(),
                            pagination.layout.as_ref(),
                            pagination.measured_with_floats,
                        ) else {
                            return Ok(None);
                        };
                        (
                            region_measurement_frames(blocks.iter(), input, &regions).0,
                            initial_float_page_geometry(input, &regions),
                            layout.pages.len(),
                        )
                    };
                    let default_width = widths.first().copied().unwrap_or(0.0);
                    if !self.lays_out_as_one_section(&regions, section_breaks(blocks))
                        || !collect_note_refs(blocks).is_empty()
                        || docx_layout::measure_blocks::has_floating_zones(
                            blocks,
                            default_width,
                            measurement.as_ref(),
                            Some(&geometry),
                        )?
                    {
                        return Ok(None);
                    }
                    match self.resident_layout_input_from_blocks(
                        blocks,
                        true,
                        self.local_lowering.get(),
                        &mut |index, _key, _previous_block, next_block| {
                            let width = widths.get(index).copied().unwrap_or(default_width);
                            resolve_line_unit_spacing(
                                next_block,
                                regions.paragraph_spacing_line_px(0),
                            );
                            resolve_doc_grid_pitch(next_block, regions.doc_grid_snap_pitch_px(0));
                            // as the region pass measures a changed table
                            docx_layout::paragraph_spacing::apply_contextual_spacing_blocks(
                                std::slice::from_mut(next_block),
                            );
                            docx_layout::measure_blocks::measure_block(
                                next_block,
                                width,
                                measurement.as_ref(),
                            )
                        },
                    ) {
                        Ok(mut resident) => {
                            resident.lowering = Some(lowering);
                            Ok(Some((resident, previous_pages, widths)))
                        }
                        Err(_) => Ok(None),
                    }
                },
            )
            .map_err(|error| error.to_string())??;
        let (resident, previous_pages, widths) = match outcome {
            Some(resident) => resident,
            None => return Ok(false),
        };
        phase(RegionResidentPhase::Measured);
        let previous_capture = self.capture.borrow_mut().take();
        if let Err(error) = self.layout_document_value_with_fingerprints(
            resident.input,
            resident.block_fingerprints,
            None,
            cached_page_totals,
        ) {
            if self.local_lowering.get() {
                // the walk moved the retained extents out
                let mut pagination = self.pagination.borrow_mut();
                pagination.block_fingerprints.clear();
                pagination.measured_with = None;
            }
            return Err(error);
        }
        let mut pagination = self.pagination.borrow_mut();
        pagination.input_lowering = resident.lowering;
        // The fast path measures through the region config too, so its
        // retained arena is also eligible for the next pass's reuse walk.
        pagination.measured_with = Some(measurement_fingerprint);
        pagination.measured_font_dependencies = resident.font_dependencies;
        pagination.measured_widths = widths;
        let serial = pagination.layout_epoch;
        let layout = pagination
            .layout
            .as_mut()
            .expect("layout retained after successful pagination");
        let mut restamped_pages = Vec::new();
        apply_document_regions_tracked(layout, &regions, &mut restamped_pages);
        let unchanged = layout.pages.len() == previous_pages;
        if let Some(pages) = &mut pagination.restamped_pages {
            pages.extend(restamped_pages);
        }
        drop(pagination);
        if unchanged && let Some(capture) = previous_capture {
            self.capture.replace(Some(LayoutCapture {
                version: self.doc.version(),
                serial,
                ..capture
            }));
        }
        Ok(unchanged)
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
        let shown = if let Some(headers_footers) = headers_footers {
            fields.insert("headersFooters".to_owned(), headers_footers)
        } else {
            fields.remove("headersFooters")
        };
        let rebuilt = serde_json::to_string(&value)
            .map_err(|error| format!("serialize display extras: {error}"))?;
        if rebuilt == extras {
            return Ok(extras);
        }
        // The host writes the shown frame's extras in its own JSON. When they
        // carry these headers and footers, keep them verbatim, so the next
        // frame builds on the shown one.
        let reparsed: serde_json::Value = serde_json::from_str(&rebuilt)
            .map_err(|error| format!("parse display extras: {error}"))?;
        if json_option_equal(reparsed.get("headersFooters"), shown.as_ref()) {
            return Ok(extras);
        }
        Ok(rebuilt)
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
                None,
                false,
            )?;
            self.pagination.borrow_mut().input_lowering = resident.lowering;
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
        self.pagination.borrow_mut().restamped_pages = Some(BTreeSet::new());
        Ok(display_json)
    }

    /// Build the retained display list and return a binary FrameDelta v1.
    /// `expected_frame_epoch` is the last frame the host actually applied. A
    /// mismatch automatically widens to a full recovery frame, and the new
    /// frame's epoch always exceeds it, so a host switching engines can apply it.
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
        let (incremental_build, rebuilt_display_pages, rebuilt_pages, shifts) = {
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
            let caret = if display.windowed_incremental_builds && display.window.is_some() {
                self.resident_caret_head
                    .borrow()
                    .as_ref()
                    .map(|(story, head)| {
                        let extent = || {
                            let txn = self.doc.yrs_doc().transact();
                            let index = head.get_offset(&txn)?.index;
                            drop(txn);
                            let paragraphs = self.doc.paragraph_index(story).ok()?;
                            let paragraph = paragraphs.para_at(index)?;
                            let (epoch, map) = pagination.input_lowering.as_ref()?;
                            if *epoch != self.doc_epoch() {
                                return None;
                            }
                            lowered_caret_position(map, input, story, paragraph, index)
                                .map(CaretExtent::Position)
                        };
                        extent().unwrap_or(CaretExtent::Unmapped)
                    })
            } else {
                None
            };
            let font_cache_identity =
                docx_layout::measure_font_cache_identity(&display.font_chains);
            let build = if pagination.last_incremental
                && display.extras_fingerprint == extras_fingerprint
                && display.font_cache_identity == Some(font_cache_identity)
            {
                // The first range is rebuilt as a range; the pages after it shift,
                // but later ranges, the pages elsewhere whose notes anchor to
                // references the edit moved, and retained pages whose section or
                // numbering stamps changed are rebuilt too.
                let first = pagination
                    .rebuilt_page_ranges
                    .first()
                    .cloned()
                    .unwrap_or(pagination.rebuilt_page_start..pagination.rebuilt_page_end);
                let restamped = match &pagination.restamped_pages {
                    Some(pages) => pages
                        .iter()
                        .copied()
                        .filter(|&index| index < layout.pages.len())
                        .collect::<Vec<_>>(),
                    None => display
                        .list
                        .as_ref()
                        .filter(|list| list.pages.len() == layout.pages.len())
                        .map(|list| {
                            list.pages
                                .iter()
                                .zip(&layout.pages)
                                .enumerate()
                                .filter(|(_, (shown, page))| !page_stamps_match(shown, page))
                                .map(|(index, _)| index)
                                .collect::<Vec<_>>()
                        })
                        .unwrap_or_default(),
                };
                let note_pages: Vec<usize> = pagination
                    .rebuilt_page_ranges
                    .iter()
                    .skip(1)
                    .flat_map(Clone::clone)
                    .chain(pagination.note_changed_pages.iter().copied())
                    .chain(restamped)
                    .filter(|&index| !first.contains(&index))
                    .collect::<BTreeSet<_>>()
                    .into_iter()
                    .collect();
                let rebuilt_pages: HashSet<usize> =
                    first.clone().chain(note_pages.iter().copied()).collect();
                let build =
                    window_build_pages(&display, layout, rebuilt_pages.iter().copied(), caret);
                let shifts = if let DisplayState {
                    list: Some(previous),
                    resident_input: Some(resident_input),
                    ..
                } = &mut *display
                {
                    docx_layout::update_resident_display_list_incremental_partial_shifts_observed(
                        input,
                        layout,
                        resident_input,
                        previous,
                        first.start,
                        first.end,
                        &note_pages,
                        &pagination.position_deltas,
                        &|index| build.as_ref().is_none_or(|pages| pages.contains(&index)),
                        observe_display_phase,
                    )?
                } else {
                    None
                };
                let incremental = shifts.is_some();
                if !incremental {
                    let build = windowed_full_build_pages(&display, layout, caret);
                    let (resident_input, list) =
                        docx_layout::build_resident_display_list_partial_observed(
                            input,
                            layout,
                            extras_json,
                            &|index| build.as_ref().is_none_or(|pages| pages.contains(&index)),
                            observe_display_phase,
                        )?;
                    display.resident_input = Some(resident_input);
                    display.list = Some(list);
                }
                let rebuilt_display_pages = if incremental {
                    rebuilt_pages
                        .iter()
                        .filter(|&&index| build.as_ref().is_none_or(|pages| pages.contains(&index)))
                        .count()
                } else {
                    rebuilt_pages.len()
                };
                (
                    incremental,
                    rebuilt_display_pages,
                    rebuilt_pages,
                    shifts.unwrap_or_default(),
                )
            } else {
                let build = windowed_full_build_pages(&display, layout, caret);
                let (resident_input, list) =
                    docx_layout::build_resident_display_list_partial_observed(
                        input,
                        layout,
                        extras_json,
                        &|index| build.as_ref().is_none_or(|pages| pages.contains(&index)),
                        observe_display_phase,
                    )?;
                display.resident_input = Some(resident_input);
                display.list = Some(list);
                (
                    false,
                    layout.pages.len(),
                    HashSet::new(),
                    docx_layout::display_list::IncrementalDisplayShifts::default(),
                )
            };
            if display.extras_json.is_none() || display.extras_fingerprint != extras_fingerprint {
                let chains = display.resident_input.as_ref().map(|input| {
                    input
                        .font_chains()
                        .iter()
                        .map(|(key, ids)| (key.clone(), ids.clone()))
                        .collect()
                });
                if let Some(chains) = chains {
                    display.font_chains = chains;
                }
            }
            build
        };
        self.pagination.borrow_mut().restamped_pages = Some(BTreeSet::new());
        observe_display_phase();
        let mut display = self.display.borrow_mut();
        display.frame_epoch = display
            .frame_epoch
            .max(expected_frame_epoch)
            .wrapping_add(1);
        display.display_builds = display.display_builds.wrapping_add(1);
        display.incremental_display_builds = display
            .incremental_display_builds
            .wrapping_add(u64::from(incremental_build));
        display.rebuilt_display_pages = display
            .rebuilt_display_pages
            .wrapping_add(rebuilt_display_pages as u64);
        display.extras_fingerprint = extras_fingerprint;
        display.font_cache_identity = Some(docx_layout::measure_font_cache_identity(
            &display.font_chains,
        ));
        display.extras_json = Some(extras_json.to_owned());
        let frame_epoch = display.frame_epoch;
        let binary_frame_epoch = display.binary_frame_epoch;
        let full = display.fresh_base
            || expected_frame_epoch != binary_frame_epoch
            || binary_frame_epoch == 0;
        let layout_epoch = self.pagination.borrow().layout_epoch;
        // Split borrows: the encoder reads the retained list and the previous
        // snapshots in place — no per-frame deep clone of the snapshot set.
        let display = &mut *display;
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
        let bytes = if incremental_build && !full && display.pages.len() == list.pages.len() {
            let mut rebuilt: Vec<_> = rebuilt_pages.into_iter().collect();
            rebuilt.sort_unstable();
            let runs: Vec<_> = shifts
                .runs
                .iter()
                .map(|run| PageShiftRun {
                    start: run.start,
                    end: run.end,
                    delta: run.delta,
                })
                .collect();
            encode_frame_delta_changes(
                list,
                &mut display.pages,
                epochs,
                DisplayChanges {
                    rebuilt: &rebuilt,
                    repositioned: &shifts.mixed,
                    shifts: &runs,
                },
            )?
        } else {
            for snapshot in &mut display.pages {
                snapshot.materialize_positions();
            }
            let mut next_page_id = display.next_page_id;
            let (bytes, pages) =
                encode_frame_delta(list, &display.pages, epochs, full, &mut next_page_id)?;
            display.pages = pages;
            display.next_page_id = next_page_id;
            bytes
        };
        display.binary_frame_epoch = frame_epoch;
        display.encoded_doc_epoch = epochs.doc_epoch;
        display.encoded_layout_epoch = epochs.layout_epoch;
        display.fresh_base = false;
        Ok(bytes)
    }

    /// Limit full builds to `window` and previously built pages; `None` builds all.
    pub fn set_display_window(&self, window: Option<std::ops::Range<usize>>) {
        self.display.borrow_mut().window = window;
    }

    /// While set, windowed builds keep every page the previous list had built,
    /// as with windowed builds off.
    pub fn set_display_retain_built_pages(&self, retain: bool) {
        self.display.borrow_mut().retain_built_pages = retain;
    }

    /// Limit incremental rebuilds to the display window and caret pages. Off by default.
    pub fn set_windowed_incremental_builds(&self, enabled: bool) {
        self.display.borrow_mut().windowed_incremental_builds = enabled;
    }

    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn set_resident_caret_head(&self, head: Option<(String, StickyIndex)>) {
        *self.resident_caret_head.borrow_mut() = head;
    }

    /// Build the requested pages that are still unbuilt and return a
    /// FrameDelta v1 carrying them. `expected_frame_epoch` works as for
    /// [`Self::build_display_list_frame`].
    pub fn build_display_pages_frame(
        &self,
        pages: &[usize],
        expected_frame_epoch: u64,
    ) -> Result<Vec<u8>, String> {
        let built = {
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
            let DisplayState {
                list: Some(list),
                resident_input: Some(resident_input),
                ..
            } = &mut *display
            else {
                return Err("resident display list is not built".to_owned());
            };
            docx_layout::build_resident_display_pages(input, layout, resident_input, list, pages)?
        };
        self.encode_display_pages_frame(&built, expected_frame_epoch)
    }

    /// Release display pages; an empty result means the request was superseded.
    pub fn release_display_pages_frame(
        &self,
        pages: &[usize],
        expected_frame_epoch: u64,
    ) -> Result<Vec<u8>, String> {
        let released = {
            let pagination = self.pagination.borrow();
            let mut display = self.display.borrow_mut();
            if expected_frame_epoch != display.binary_frame_epoch
                || display.binary_frame_epoch == 0
                || display.frame_epoch != display.binary_frame_epoch
                || display.fresh_base
                || display.encoded_doc_epoch != self.doc_epoch()
                || display.encoded_layout_epoch != pagination.layout_epoch
            {
                return Ok(Vec::new());
            }
            let input = pagination
                .input
                .as_ref()
                .ok_or_else(|| "resident pagination input is not built".to_owned())?;
            let layout = pagination
                .layout
                .as_ref()
                .ok_or_else(|| "resident layout is not built".to_owned())?;
            if display.pages.len() != layout.pages.len() {
                return Err("resident frame does not match the layout".to_owned());
            }
            let DisplayState {
                list: Some(list),
                resident_input: Some(resident_input),
                ..
            } = &mut *display
            else {
                return Err("resident display list is not built".to_owned());
            };
            docx_layout::release_resident_display_pages(input, layout, resident_input, list, pages)?
        };
        self.encode_display_pages_frame(&released, expected_frame_epoch)
    }

    fn encode_display_pages_frame(
        &self,
        changed_pages: &[usize],
        expected_frame_epoch: u64,
    ) -> Result<Vec<u8>, String> {
        let mut display = self.display.borrow_mut();
        display.frame_epoch = display
            .frame_epoch
            .max(expected_frame_epoch)
            .wrapping_add(1);
        let frame_epoch = display.frame_epoch;
        let binary_frame_epoch = display.binary_frame_epoch;
        let full = display.fresh_base
            || expected_frame_epoch != binary_frame_epoch
            || binary_frame_epoch == 0;
        let epochs = FrameEpochs {
            doc_epoch: self.doc_epoch(),
            layout_epoch: self.pagination.borrow().layout_epoch,
            frame_epoch,
            base_frame_epoch: binary_frame_epoch,
        };
        let display = &mut *display;
        let list = display
            .list
            .as_ref()
            .expect("display list built before FrameDelta encoding");
        let bytes = if !full && display.pages.len() == list.pages.len() {
            encode_frame_delta_changes(
                list,
                &mut display.pages,
                epochs,
                DisplayChanges {
                    rebuilt: changed_pages,
                    repositioned: &[],
                    shifts: &[],
                },
            )?
        } else {
            for snapshot in &mut display.pages {
                snapshot.materialize_positions();
            }
            let mut next_page_id = display.next_page_id;
            let (bytes, snapshots) =
                encode_frame_delta(list, &display.pages, epochs, full, &mut next_page_id)?;
            display.pages = snapshots;
            display.next_page_id = next_page_id;
            bytes
        };
        display.binary_frame_epoch = frame_epoch;
        display.encoded_doc_epoch = epochs.doc_epoch;
        display.encoded_layout_epoch = epochs.layout_epoch;
        display.fresh_base = false;
        Ok(bytes)
    }

    /// Make the next frame a full one: its caller holds another engine's frames.
    pub fn reset_frame_base(&self) {
        self.display.borrow_mut().fresh_base = true;
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
            let pagination = self.pagination.borrow();
            resident_paragraph_position(pagination.input.as_ref()?, para_id, offset)
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

    /// Exports the committed state with the page map of the retained region layout, as
    /// [`EditingDoc::export_structured`] exports it. The layout must have lowered this document
    /// version from the session's own stories with section, settings and note metadata that
    /// describe it, measured every font requirement with a registered font in the current font
    /// store, and settled its notes; nothing is laid out, flushed or changed.
    pub fn export_structured_with_pages(
        &self,
        options: &PageExportOptions,
    ) -> Result<ExportRead<DocxPagedStructuredContent<DocxLayoutMap>>, ExportRefusal> {
        self.export_with_pages(options, AnchorScope::Session, None)
    }

    /// [`Self::export_structured_with_pages`] for an editor that owns the layout inputs:
    /// `current_request_json` is the region layout request it would lay the document out with
    /// now, and the retained layout must have used the same fonts, fallback order, measurement
    /// defaults, render environment and pagination options. The editor owns the final section
    /// and settings, so they are checked against this request rather than the source package.
    pub fn export_structured_with_pages_for(
        &self,
        options: &PageExportOptions,
        current_request_json: &str,
    ) -> Result<ExportRead<DocxPagedStructuredContent<DocxLayoutMap>>, ExportRefusal> {
        self.export_with_pages(options, AnchorScope::Session, Some(current_request_json))
    }

    /// [`Self::export_structured_with_pages`] for a private session opened from DOCX bytes:
    /// anchors address the returned content and the map carries no session token.
    pub fn export_snapshot_with_pages(
        &self,
        options: &PageExportOptions,
    ) -> Result<DocxPagedStructuredContent<DocxSnapshotLayoutMap>, ExportFailure> {
        let read = self
            .export_with_pages(options, AnchorScope::Snapshot, None)
            .map_err(|refusal| refusal.failure)?;
        Ok(DocxPagedStructuredContent {
            structured: read.content.structured,
            layout: DocxSnapshotLayoutMap::from_session(read.content.layout),
        })
    }

    /// Lays this private session out with `fonts` alone and exports it as
    /// [`Self::export_snapshot_with_pages`] does. The fonts are registered, in order, in a
    /// measurement font store of their own, so `request_json`'s font chains name them by their
    /// index, and the store in use is left exactly as it was. The outer error is a font
    /// the engine rejects or a request it cannot lay out.
    pub fn export_snapshot_with_private_fonts(
        &self,
        fonts: &[&[u8]],
        request_json: &str,
        options: &PageExportOptions,
    ) -> Result<Result<DocxPagedStructuredContent<DocxSnapshotLayoutMap>, ExportFailure>, String>
    {
        docx_layout::with_private_measure_fonts(|| {
            for (index, bytes) in fonts.iter().enumerate() {
                let id = docx_layout::register_measure_font_bytes(bytes)
                    .map_err(|error| format!("font {index}: {error}"))?;
                if id as usize != index {
                    return Err(format!("font {index} was registered as font {id}"));
                }
            }
            self.layout_document_with_regions_value(request_json)?;
            Ok(self.export_snapshot_with_pages(options))
        })
    }

    fn export_with_pages(
        &self,
        options: &PageExportOptions,
        scope: AnchorScope,
        current: Option<&str>,
    ) -> Result<ExportRead<DocxPagedStructuredContent<DocxLayoutMap>>, ExportRefusal> {
        let version = self.doc.version();
        let refuse = |code, message: &str| ExportRefusal {
            version: version.clone(),
            failure: pages::failure(code, message),
        };
        let export_options = options.export_options();
        let limits = PageLimits::new(options).map_err(|failure| ExportRefusal {
            version: version.clone(),
            failure,
        })?;
        crate::structured::validate_options(&export_options).map_err(|failure| ExportRefusal {
            version: version.clone(),
            failure,
        })?;
        let capture = self.capture.borrow();
        let regions = self.regions.borrow();
        let pagination = self.pagination.borrow();
        let (Some(capture), Some(region_state), Some(input), Some(layout)) = (
            capture.as_ref(),
            regions.as_ref(),
            pagination.input.as_ref(),
            pagination.layout.as_ref(),
        ) else {
            return Err(refuse(
                ExportFailureCode::LayoutUnavailable,
                "No complete region layout of this document is retained; lay it out first.",
            ));
        };
        if capture.version != version {
            return Err(refuse(
                ExportFailureCode::StaleDocument,
                "The document changed after it was laid out; lay it out again.",
            ));
        }
        if capture.fonts != docx_layout::measure_fonts_generation() {
            return Err(refuse(
                ExportFailureCode::StaleLayout,
                "The measurement fonts changed after the layout; lay it out again.",
            ));
        }
        let layout_version = format!("{version}:{}", capture.serial);
        if options
            .expect_layout_version
            .as_ref()
            .is_some_and(|expected| *expected != layout_version)
        {
            return Err(refuse(
                ExportFailureCode::StaleLayout,
                "The layout is not the one expectLayoutVersion names.",
            ));
        }
        let note_fallback_pages: &[usize] = match &capture.note_settlement {
            NoteSettlement::Converged => &[],
            NoteSettlement::Covering(pages) => pages,
            NoteSettlement::Unsettled => {
                return Err(refuse(
                    ExportFailureCode::LayoutNotConverged,
                    "Note placement did not settle in the retained layout.",
                ));
            }
        };
        if !capture.render_env.revision_preview.is_empty() {
            return Err(refuse(
                ExportFailureCode::UnsupportedRevisionLayout,
                "The retained layout previews revision decisions instead of their markup.",
            ));
        }
        let env = &capture.render_env;
        let request: serde_json::Value =
            serde_json::from_str(&region_state.request_json).map_err(|error| {
                refuse(
                    ExportFailureCode::LayoutUnavailable,
                    &format!("The retained layout request cannot be read: {error}"),
                )
            })?;
        let current = current
            .map(serde_json::from_str::<serde_json::Value>)
            .transpose()
            .map_err(|error| {
                refuse(
                    ExportFailureCode::LayoutUnavailable,
                    &format!("The editor's layout request cannot be read: {error}"),
                )
            })?;
        if current
            .as_ref()
            .and_then(|request| request.get("renderEnv"))
            .and_then(|env| env.get("revisionPreview"))
            .is_some_and(|preview| !RenderEnv::parse_revision_preview(preview).is_empty())
        {
            return Err(refuse(
                ExportFailureCode::UnsupportedRevisionLayout,
                "The retained layout previews revision decisions instead of their markup.",
            ));
        }
        if let Some(message) = pages::metadata_mismatch(&self.doc, &request, current.is_some()) {
            return Err(refuse(ExportFailureCode::StaleLayout, &message));
        }
        let laid_out: HashSet<i64> = capture
            .notes
            .iter()
            .map(|content| content.map_id())
            .collect();
        if let Some(reference) = input
            .measured
            .iter()
            .flat_map(|measured| collect_note_refs(std::slice::from_ref(&measured.block)))
            .find(|reference| !laid_out.contains(&reference.map_id()))
        {
            return Err(refuse(
                ExportFailureCode::StaleLayout,
                &format!(
                    "The layout's note metadata has no {} {}, which the document references; lay it out with its current notes.",
                    match reference.note_kind {
                        docx_layout::footnotes::NoteKind::Footnote => "footnote",
                        docx_layout::footnotes::NoteKind::Endnote => "endnote",
                    },
                    reference.note_id
                ),
            ));
        }
        let measurement = request.get("measurement").cloned().unwrap_or_default();
        let blocks = input
            .measured
            .iter()
            .map(|measured| &measured.block)
            .chain(
                capture
                    .headers_footers
                    .iter()
                    .flat_map(|payload| &payload.variants)
                    .flat_map(|variant| variant.measured.iter().map(|measured| &measured.block)),
            )
            .chain(capture.notes.iter().flat_map(|content| &content.blocks));
        let missing = missing_font_chains(&measurement, blocks);
        if !missing.is_empty() {
            return Err(refuse(
                ExportFailureCode::LayoutUnavailable,
                &format!(
                    "No registered font measures {}; lay the document out with a font for every font it uses.",
                    missing
                        .iter()
                        .take(5)
                        .map(|key| format!("{key:?}"))
                        .collect::<Vec<_>>()
                        .join(", ")
                ),
            ));
        }
        let extents = input
            .measured
            .iter()
            .map(|measured| &measured.measure)
            .chain(
                capture
                    .headers_footers
                    .iter()
                    .flat_map(|payload| &payload.variants)
                    .flat_map(|variant| variant.measured.iter().map(|measured| &measured.measure)),
            )
            .chain(capture.notes.iter().flat_map(|content| &content.measures));
        if docx_layout::measure_blocks::measured_synthetically(extents) {
            return Err(refuse(
                ExportFailureCode::LayoutUnavailable,
                "Some text was laid out with stand-in metrics because its fonts could not measure it; lay it out again with fonts that can.",
            ));
        }
        let Some(font_set_fingerprint) = self.font_set_fingerprint(&request) else {
            return Err(refuse(
                ExportFailureCode::LayoutUnavailable,
                "A font the layout measured with is no longer registered.",
            ));
        };
        let options_fingerprint = layout_options_fingerprint(request.clone());
        if let Some(current) = current
            && (self.font_set_fingerprint(&current).as_ref() != Some(&font_set_fingerprint)
                || layout_options_fingerprint(current) != options_fingerprint)
        {
            return Err(refuse(
                ExportFailureCode::StaleLayout,
                "The editor's fonts, measurement defaults, render environment or pagination options differ from the layout's; lay it out again.",
            ));
        }
        let mut roots = vec!["body".to_owned()];
        if let Some(payload) = capture.headers_footers.as_deref() {
            roots.extend(
                payload
                    .variants
                    .iter()
                    .map(|variant| format!("hf:{}", variant.r_id)),
            );
        }
        roots.extend(
            capture
                .notes
                .iter()
                .filter(|content| !content.blocks.is_empty())
                .map(|content| {
                    let prefix = match content.note_kind {
                        docx_layout::footnotes::NoteKind::Footnote => "fn",
                        docx_layout::footnotes::NoteKind::Endnote => "en",
                    };
                    format!("{prefix}:{}", content.id)
                }),
        );
        let mut maps = HashMap::new();
        for root in roots {
            if maps.contains_key(&root) {
                continue;
            }
            let map = self.lowering_map(&root, env).map_err(|error| {
                refuse(
                    ExportFailureCode::LayoutUnavailable,
                    &format!("Story {root} can no longer be lowered: {error}"),
                )
            })?;
            maps.insert(root, map);
        }
        if options.revision_view != RevisionView::Markup {
            let stories: BTreeSet<String> = maps
                .values()
                .flat_map(|map| map.stories.iter().cloned())
                .collect();
            if let Some(story) = pages::revised_story(&self.doc, &stories) {
                return Err(refuse(
                    ExportFailureCode::UnsupportedRevisionLayout,
                    &format!(
                        "Story {story} has pending revisions, so the pages, laid out with revision markup, do not show the {} view; export the markup view.",
                        match options.revision_view {
                            RevisionView::Accepted => "accepted",
                            _ => "original",
                        }
                    ),
                ));
            }
        }
        let read = self.doc.export_structured_scoped(&export_options, scope)?;
        for story in &read.content.stories {
            if matches!(story.kind, StoryKind::Header | StoryKind::Footer)
                && !maps.contains_key(&story.story)
                && let Ok(map) = self.lowering_map(&story.story, env)
            {
                maps.insert(story.story.clone(), map);
            }
        }
        let display = limits
            .include_geometry
            .then(|| {
                let extras = serde_json::json!({
                    "contractVersion": input.options.contract_version,
                    "headersFooters": region_state.headers_footers,
                    "fontChains": request.get("measurement").and_then(|value| value.get("fontChains")),
                });
                docx_layout::build_display_list_value_from_resident(
                    input,
                    layout,
                    &extras.to_string(),
                )
                .ok()
            })
            .flatten();
        let map = pages::build_layout_map(
            &self.doc,
            &read.content,
            &pages::CapturedLayout {
                layout,
                measured: &input.measured,
                headers_footers: capture.headers_footers.as_deref(),
                bands_composed: region_state.headers_footers.is_some(),
                notes: &capture.notes,
                note_fallback_pages,
                maps: &maps,
                display: display.as_ref(),
            },
            &limits,
            pages::MapIdentity {
                document_version: version.clone(),
                layout_version,
                layout_epoch: capture.serial.to_string(),
                font_set_fingerprint,
                options_fingerprint,
            },
        )
        .map_err(|failure| ExportRefusal {
            version: version.clone(),
            failure,
        })?;
        Ok(ExportRead {
            version: read.version,
            content: DocxPagedStructuredContent {
                structured: read.content,
                layout: map,
            },
        })
    }

    /// The lowering map of `story` for this document epoch and environment.
    fn lowering_map(&self, story: &str, env: &RenderEnv) -> Result<Rc<LoweringMap>, BridgeError> {
        let epoch = self.doc_epoch();
        if !self.story_is_resident(story, epoch, env) {
            self.lower_story_into_cache(story, epoch, env)?;
        }
        Ok(Rc::clone(
            &self
                .render
                .borrow()
                .stories
                .get(story)
                .expect("resident story exists after lowering")
                .map,
        ))
    }

    /// A fingerprint of the fonts, fallback order and measurement defaults of `request`, by font
    /// content rather than store ids. `None` when a font it names is not registered.
    fn font_set_fingerprint(&self, request: &serde_json::Value) -> Option<String> {
        let measurement = request.get("measurement").cloned().unwrap_or_default();
        let store = docx_layout::measure_store_id();
        let mut chains = BTreeMap::new();
        if let Some(serde_json::Value::Object(font_chains)) = measurement.get("fontChains") {
            for (key, ids) in font_chains {
                let mut faces = Vec::new();
                for id in ids.as_array().into_iter().flatten() {
                    let id = u32::try_from(id.as_u64()?).ok()?;
                    let cached = self.font_fingerprints.borrow().get(&(store, id)).cloned();
                    let face = match cached {
                        Some(face) => face,
                        None => {
                            let face = docx_layout::with_measure_face(id, |bytes, metrics| {
                                pages::sha256_hex(
                                    [pages::sha256_hex(bytes).as_bytes(), metrics.as_bytes()]
                                        .concat()
                                        .as_slice(),
                                )
                            })?;
                            self.font_fingerprints
                                .borrow_mut()
                                .insert((store, id), face.clone());
                            face
                        }
                    };
                    faces.push(face);
                }
                chains.insert(key.clone(), faces);
            }
        }
        let canonical = serde_json::json!({
            "fontChains": chains,
            "defaults": measurement.get("defaults"),
            "compat": measurement.get("compat"),
            "authoritativeShaping": measurement.get("authoritativeShaping"),
        });
        Some(pages::sha256_hex(canonical_json(&canonical).as_bytes()))
    }
}

#[cfg(test)]
use crate::seed::fixture as lowering_fixture;

#[cfg(test)]
#[allow(dead_code)]
#[path = "../tests/support/page_fixture.rs"]
mod lowering_pages;

#[cfg(test)]
#[allow(dead_code)]
#[path = "../tests/support/preview_fixture.rs"]
mod preview_fixture;

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use yrs::Any;
    use yrs::types::Attrs;

    fn table_wrap_section(content_width: f64, columns: serde_json::Value) -> serde_json::Value {
        json!({
            "pageSize": {"w": content_width + 200.0, "h": 800},
            "margins": {"top": 100, "right": 100, "bottom": 100, "left": 100},
            "columns": columns
        })
    }

    fn measured_table_wrap_margins(
        anchor: &str,
        alignment: &str,
        sections: serde_json::Value,
    ) -> (f64, f64) {
        let regions: DocumentRegions =
            serde_json::from_value(json!({"sections": sections})).unwrap();
        let mut blocks: Vec<LayoutBlock> = regions
            .sections
            .iter()
            .take(regions.sections.len() - 1)
            .enumerate()
            .map(|(index, section)| {
                serde_json::from_value(json!({
                    "kind": "sectionBreak", "id": index, "type": section.section_start
                }))
                .unwrap()
            })
            .collect();
        blocks.extend(
            serde_json::from_value::<Vec<LayoutBlock>>(json!([
                {"kind": "columnBreak", "id": "column-two"},
                {
                    "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
                    "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                    "floating": {
                        "horzAnchor": anchor, "tblpXSpec": "right", "vertAnchor": "text", "tblpY": 1,
                        "leftFromText": 9, "rightFromText": 13
                    }
                },
                {
                    "kind": "paragraph", "id": "text", "attrs": {"alignment": alignment},
                    "runs": [{"kind": "text", "text": if alignment == "center" {
                        "short line".to_owned()
                    } else {
                        "Long left-aligned prose following the floating table. ".repeat(20)
                    }}]
                }
            ]))
            .unwrap(),
        );
        let mut input = LayoutInput {
            measured: Vec::new(),
            options: Default::default(),
        };
        apply_section_geometry_to_blocks(&mut blocks, &mut input.options, &regions);
        let (widths, table_wrap_frames) =
            region_measurement_frames(blocks.iter(), &input, &regions);
        assert_eq!(widths[widths.len() - 2], 600.0);
        let font = docx_layout::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let measurement = docx_layout::measure_blocks::MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily": "Liberation Sans", "fontSize": 12}),
            ..Default::default()
        };
        let geometry = initial_float_page_geometry(&input, &regions);
        let measures = docx_layout::measure_blocks::measure_blocks_with_table_wrap_frames(
            &mut blocks,
            &widths,
            &table_wrap_frames,
            &measurement,
            Some(&geometry),
            &BTreeMap::new(),
        )
        .unwrap();
        let BlockExtent::Paragraph(paragraph) = measures.last().unwrap() else {
            panic!()
        };
        let line = &paragraph.lines[0];
        assert_ne!(line.synthetic_fallback, Some(true));
        (
            line.left_offset.unwrap_or(0.0),
            line.right_offset.unwrap_or(0.0),
        )
    }

    #[test]
    fn column_anchored_wide_tables_keep_main_margins() {
        for section in [
            table_wrap_section(600.0, serde_json::Value::Null),
            table_wrap_section(1220.0, json!({"count": 2, "gap": 20})),
        ] {
            assert_eq!(
                measured_table_wrap_margins("column", "left", json!([section])),
                (0.0, 0.0)
            );
        }
    }

    #[test]
    fn wide_tables_in_unequal_columns_keep_main_margins() {
        let section = table_wrap_section(
            1020.0,
            json!({"count": 2, "gap": 20, "equalWidth": false, "columns": [{"width": 600}, {"width": 400}]}),
        );
        assert_eq!(
            measured_table_wrap_margins("text", "left", json!([section])),
            (0.0, 0.0)
        );
    }

    #[test]
    fn centered_wide_tables_in_equal_columns_keep_main_margins() {
        let section = table_wrap_section(1220.0, json!({"count": 2, "gap": 20}));
        assert_eq!(
            measured_table_wrap_margins("text", "center", json!([section])),
            (0.0, 0.0)
        );
    }

    #[test]
    fn wide_tables_in_a_known_single_column_use_the_larger_gap() {
        let section = table_wrap_section(600.0, serde_json::Value::Null);
        assert_eq!(
            measured_table_wrap_margins("text", "left", json!([section])),
            (0.0, 369.0)
        );
    }

    #[test]
    fn negative_indents_clear_the_frame_flags_of_their_whole_float_flow() {
        let regions: DocumentRegions =
            serde_json::from_value(json!({"sections": [table_wrap_section(600.0, json!(null))]}))
                .unwrap();
        for (break_before, expected) in
            [(false, [false, false, false]), (true, [true, true, false])]
        {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {
                    "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
                    "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                    "floating": {"horzAnchor": "text", "tblpX": 140}
                },
                {"kind": "paragraph", "id": "text", "runs": []},
                {"kind": "paragraph", "id": "later", "runs": [],
                 "attrs": {"indent": {"right": -100}, "pageBreakBefore": break_before}}
            ]))
            .unwrap();
            let mut input = LayoutInput {
                measured: Vec::new(),
                options: Default::default(),
            };
            apply_section_geometry_to_blocks(&mut blocks, &mut input.options, &regions);
            let (_, frames) = region_measurement_frames(blocks.iter(), &input, &regions);
            assert_eq!(frames, expected, "pageBreakBefore={break_before}");
        }
    }

    #[test]
    fn continuous_content_width_changes_keep_main_table_margins() {
        let first = table_wrap_section(600.0, serde_json::Value::Null);
        let second = table_wrap_section(500.0, serde_json::Value::Null);
        let mut third = table_wrap_section(600.0, serde_json::Value::Null);
        third["sectionStart"] = json!("continuous");
        assert_eq!(
            measured_table_wrap_margins("text", "left", json!([first, second, third])),
            (0.0, 0.0)
        );
    }

    #[test]
    fn seeded_session_epochs_advance_before_local_and_remote_update_listeners() {
        let engine = EngineSession::new(7);
        crate::seed::seed_from_docx(
            engine.doc(),
            &docx_bytes("", "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"),
        )
        .unwrap();
        let peer = EditingDoc::new(8);
        peer.apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        let events = Rc::new(RefCell::new(Vec::new()));
        let observed = Rc::clone(&events);
        let epoch = Rc::clone(&engine.doc_epoch);
        let _listener = engine
            .doc()
            .yrs_doc()
            .observe_update_v1(move |txn, event| {
                observed.borrow_mut().push((
                    epoch.get(),
                    txn.origin().is_none(),
                    event.update.clone(),
                ));
            })
            .unwrap();
        let seeded_epoch = engine.doc_epoch();
        let seeded_version = engine.doc().version();
        drop(engine.doc().yrs_doc().transact_mut());
        assert_eq!(engine.doc_epoch(), seeded_epoch);
        assert_eq!(engine.doc().version(), seeded_version);
        assert!(events.borrow().is_empty());

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 5),
                "!",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let edited_epoch = engine.doc_epoch();
        let edited_version = engine.doc().version();
        assert_eq!(edited_epoch, seeded_epoch + 1);
        assert_ne!(edited_version, seeded_version);
        assert_eq!(events.borrow().len(), 1);
        assert_eq!(events.borrow()[0].0, edited_epoch);
        assert!(!events.borrow()[0].1);
        peer.apply_update_v1(&events.borrow()[0].2).unwrap();

        let vector = engine.doc().encode_state_vector_v1();
        peer.apply_raw_ops(
            "body",
            vec![crate::RawOp::Delete { index: 0, len: 1 }],
            &crate::EditCtx::local("", ""),
        )
        .unwrap();
        let update = peer.encode_diff_v1(&vector).unwrap();
        engine.doc().apply_update_v1(&update).unwrap();
        assert_eq!(engine.doc().encode_state_vector_v1(), vector);
        assert_eq!(engine.doc_epoch(), edited_epoch + 1);
        assert_ne!(engine.doc().version(), edited_version);
        assert_eq!(events.borrow().len(), 2);
        assert_eq!(events.borrow()[1].0, engine.doc_epoch());
        assert!(events.borrow()[1].1);
        assert!(!events.borrow()[1].2.is_empty());

        let merged_version = engine.doc().version();
        engine.doc().apply_update_v1(&update).unwrap();
        assert_eq!(engine.doc_epoch(), edited_epoch + 1);
        assert_eq!(engine.doc().version(), merged_version);
        assert_eq!(events.borrow().len(), 2);
    }

    #[test]
    fn measured_fingerprints_preserve_relative_run_positions() {
        let measured = |start: f64, gap: f64| {
            serde_json::from_value::<MeasuredBlock>(serde_json::json!({
                "block": {
                    "kind": "paragraph", "id": "joined", "pmStart": start, "pmEnd": start + 12.0,
                    "runs": [
                        {"kind": "text", "text": "before", "pmStart": start + 1.0, "pmEnd": start + 7.0},
                        {"kind": "text", "text": "yes", "pmStart": start + gap, "pmEnd": start + gap + 3.0}
                    ]
                },
                "measure": {"kind": "paragraph", "lines": [], "totalHeight": 0.0}
            }))
            .unwrap()
        };
        let original = measured_fingerprint(&measured(0.0, 8.0)).unwrap();
        assert_eq!(
            original,
            measured_fingerprint(&measured(100.0, 8.0)).unwrap()
        );
        assert_ne!(original, measured_fingerprint(&measured(0.0, 7.0)).unwrap());
    }

    #[test]
    fn measured_fingerprints_tell_anchored_images_apart() {
        let measured = |anchored: bool| {
            serde_json::from_value::<MeasuredBlock>(serde_json::json!({
                "block": {
                    "kind": "image", "id": "i", "src": "", "width": 20, "height": 20,
                    "anchor": {"isAnchored": anchored},
                    "effects": [null, 13_100_772_350_407_709_573_u64, 18_232_552_688_281_235_959_u64]
                },
                "measure": {"kind": "image", "width": 20, "height": 20}
            }))
            .unwrap()
        };
        assert_ne!(
            measured_fingerprint(&measured(false)).unwrap(),
            measured_fingerprint(&measured(true)).unwrap()
        );
    }

    #[test]
    fn measured_block_fingerprints_match_exactly_when_the_measured_blocks_do() {
        use std::path::{Path, PathBuf};

        use serde_json::Value;

        fn collect_docx(directory: &Path, paths: &mut Vec<PathBuf>) {
            for entry in std::fs::read_dir(directory).unwrap() {
                let entry = entry.unwrap();
                let path = entry.path();
                if entry.file_type().unwrap().is_dir() {
                    collect_docx(&path, paths);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension == "docx")
                {
                    paths.push(path);
                }
            }
        }

        fn region_request(
            engine: &EngineSession,
            bytes: &[u8],
            font: u32,
        ) -> Result<String, String> {
            let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
                .map_err(|error| error.to_string())?
                .document
                .package;
            let mut sections: Vec<Value> = package
                .document
                .sections
                .clone()
                .unwrap_or_default()
                .into_iter()
                .map(|section| json!({"sectionId": section.id, "properties": section.properties}))
                .collect();
            let last = sections
                .last()
                .map(|section| section["properties"].clone())
                .unwrap_or_else(|| json!(package.document.final_section_properties));
            sections.push(json!({"properties": last}));
            let mut contents = Vec::new();
            for (notes, kind) in [
                (&package.footnotes, "footnote"),
                (&package.endnotes, "endnote"),
            ] {
                for note in notes.iter().flatten() {
                    if note.note_type.is_empty() || note.note_type == "normal" {
                        contents.push(json!({"id": note.id as i64, "noteKind": kind, "height": 0}));
                    }
                }
            }
            let mut request = json!({
                "bodyStory": "body",
                "renderEnv": {},
                "options": {"pageGap": 24},
                "regions": {"sections": sections, "settings": package.settings},
                "notes": {"contents": contents},
            });
            let requirements: Vec<Value> =
                serde_json::from_str(&engine.layout_font_requirements_json(&request.to_string())?)
                    .map_err(|error| error.to_string())?;
            let chains: serde_json::Map<String, Value> = requirements
                .iter()
                .map(|requirement| {
                    (
                        requirement["key"].as_str().unwrap().to_owned(),
                        json!([font]),
                    )
                })
                .collect();
            request["measurement"] = json!({
                "fontChains": chains,
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "compat": {"noLeading": false, "doNotExpandShiftReturn": false},
                "authoritativeShaping": true,
            });
            Ok(request.to_string())
        }

        fn remove_positions(value: &mut Value) {
            match value {
                Value::Object(object) => {
                    for key in ["pmStart", "pmEnd", "docStart", "docEnd"] {
                        object.remove(key);
                    }
                    for value in object.values_mut() {
                        remove_positions(value);
                    }
                }
                Value::Array(array) => {
                    for value in array {
                        remove_positions(value);
                    }
                }
                _ => {}
            }
        }

        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let mut paths = Vec::new();
        collect_docx(&manifest.join("tests/fixtures"), &mut paths);
        collect_docx(
            &manifest.join("../betteroffice-docx/tests/corpus/fixtures"),
            &mut paths,
        );
        paths.sort();
        let mut by_key: HashMap<String, Fingerprint> = HashMap::new();
        let mut by_fingerprint: HashMap<Fingerprint, String> = HashMap::new();
        let mut laid_out = 0;
        for path in paths {
            let bytes = std::fs::read(&path).unwrap();
            let engine = EngineSession::new(7);
            let result = (|| -> Result<(), String> {
                crate::seed::seed_from_docx(engine.doc(), &bytes)?;
                let request = region_request(&engine, &bytes, font)?;
                engine.layout_document_with_regions_retained_json(&request)?;
                Ok(())
            })();
            if let Err(error) = result {
                eprintln!("skip {}: {error}", path.display());
                continue;
            }
            laid_out += 1;
            let pagination = engine.pagination.borrow();
            let measured = &pagination.input.as_ref().unwrap().measured;
            assert_eq!(
                measured.len(),
                pagination.block_fingerprints.len(),
                "{}",
                path.display()
            );
            for (measured, &fingerprint) in measured.iter().zip(&pagination.block_fingerprints) {
                let mut value = serde_json::to_value((
                    measured,
                    relative_run_position_fingerprint(&measured.block),
                ))
                .unwrap();
                remove_positions(&mut value);
                let key = serde_json::to_string(&value).unwrap();
                if let Some(previous) = by_key.insert(key.clone(), fingerprint) {
                    assert_eq!(
                        previous,
                        fingerprint,
                        "equal measured blocks: {}",
                        path.display()
                    );
                }
                if let Some(previous) = by_fingerprint.insert(fingerprint, key.clone()) {
                    assert_eq!(previous, key, "equal fingerprints: {}", path.display());
                }
            }
        }
        assert!(laid_out >= 20, "only {laid_out} documents laid out");
        assert!(
            by_key.len() > 500,
            "only {} distinct measured blocks",
            by_key.len()
        );
    }

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
                lower_preview_patches: 0,
                lower_preview_fallbacks: 0,
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
        assert_eq!(engine.stats().pagination_calls, 1);
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

        let retained = engine.retained_headers_footers_json().unwrap().unwrap();
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(&retained).unwrap(),
            value["headersFooters"]
        );
        engine
            .layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        assert_eq!(
            engine.retained_headers_footers_json().unwrap(),
            Some(retained)
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
        for (repeated_final, local_lowering) in [(false, false), (true, false), (true, true)] {
            resident_region_fast_path_matches_the_full_pass_in(repeated_final, local_lowering);
        }
    }

    fn resident_region_fast_path_matches_the_full_pass_in(
        repeated_final: bool,
        local_lowering: bool,
    ) {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(138);
        engine.set_local_lowering(local_lowering);
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
        let mut request: serde_json::Value =
            serde_json::from_str(&small_page_request(font_id)).unwrap();
        if repeated_final {
            repeat_final_section(&mut request);
        }
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let fast = !repeated_final || local_lowering;
        let armed = engine
            .regions
            .borrow()
            .as_ref()
            .unwrap()
            .fast_path
            .is_some();
        assert_eq!(
            armed, fast,
            "repeated final {repeated_final}, local lowering {local_lowering}"
        );
        engine
            .build_display_list_frame(
                &serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0,
            )
            .unwrap();
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", 2, 2),
                Some("xx"),
                local_lowering,
            )
            .unwrap();

        let stats_before = engine.stats();
        let resident = engine
            .apply_and_layout_regions_resident("body", &mut |_| {})
            .unwrap();
        assert_eq!(resident, fast, "the region fast path absorbs the edit");
        if !resident {
            engine.apply_and_layout_regions_full().unwrap();
        }
        let stats_after = engine.stats();
        if fast {
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
        }

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
        if repeated_final && local_lowering {
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", 1),
                    "y",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            assert!(
                !engine
                    .apply_and_layout_regions_resident("body", &mut |_| {})
                    .unwrap(),
                "an edit that needs the full lowering takes the full pass"
            );
        }
    }

    #[test]
    fn resident_region_fast_path_beside_contextual_spacing_matches_a_cold_pass() {
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(lowering_pages::FONT).unwrap();
        let mut request: serde_json::Value =
            serde_json::from_str(&small_page_request(font)).unwrap();
        request["regions"]["sections"][0]["properties"]["pageHeight"] = 5760.into();
        let request = request.to_string();
        let tops = |engine: &EngineSession| -> Vec<f64> {
            engine.pagination.borrow().layout.as_ref().unwrap().pages[0]
                .fragments
                .iter()
                .filter_map(|fragment| match fragment {
                    Fragment::Paragraph(fragment) if fragment.from_line == 0 => Some(fragment.y),
                    _ => None,
                })
                .collect()
        };
        let laid_out = |client_id, contextual: bool| {
            let engine = paragraphs_engine(client_id, 3);
            engine.set_local_lowering(true);
            for paragraph in engine.doc().paragraphs("body").unwrap().iter().take(2) {
                let doc = engine.doc();
                if contextual {
                    doc.set_paragraph_attr(
                        &paragraph.para_id,
                        "contextualSpacing",
                        yrs::Any::Bool(true),
                    )
                    .unwrap();
                }
                doc.set_paragraph_attr(&paragraph.para_id, "spaceAfter", yrs::Any::Number(300.0))
                    .unwrap();
            }
            engine.layout_document_with_regions_json(&request).unwrap();
            engine.build_display_list_frame("{}", 0).unwrap();
            engine
        };
        let apart = tops(&laid_out(9606, false));
        let engine = laid_out(9605, true);
        let collapsed = tops(&engine);
        assert_eq!((apart.len(), collapsed.len()), (3, 3));
        for (index, dropped) in [0.0, 20.0, 40.0].into_iter().enumerate() {
            assert!(
                (apart[index] - collapsed[index] - dropped).abs() < 1e-6,
                "same-style paragraphs drop their contextual spacing: {apart:?} vs {collapsed:?}"
            );
        }
        let matches_cold = |label: &str| {
            assert!(
                engine
                    .apply_and_layout_regions_resident("body", &mut |_| {})
                    .unwrap(),
                "{label}: the region fast path absorbs the edit"
            );
            assert_region_state_matches_cold(&engine, &request, label);
        };
        let ctx = crate::EditCtx::local("", "");
        let longer = " and then the lazy dog wakes up, stretches and chases the fox";
        for (at, text) in [(0, Some("x")), (3, Some("y")), (3, Some(longer)), (0, None)] {
            let before = tops(&engine);
            match text {
                Some(text) => engine
                    .doc()
                    .insert_text(
                        &ctx,
                        crate::Position::new("body", at),
                        text,
                        crate::FormatPolicy::Inherit,
                    )
                    .map(|_| ()),
                None => engine
                    .doc()
                    .delete_range(&ctx, crate::StoryRange::new("body", at, at + 1))
                    .map(|_| ()),
            }
            .unwrap();
            matches_cold(&format!("[{at}] {text:?}"));
            if text == Some(longer) {
                let after = tops(&engine);
                assert!(
                    after[0] == before[0] && after[1] > before[1] && after[2] > before[2],
                    "the paragraphs after a taller edited paragraph move down: {before:?} -> {after:?}"
                );
            }
        }
        let first = engine.doc().paragraphs("body").unwrap()[0]
            .text
            .chars()
            .count() as u32;
        engine
            .doc()
            .delete_range(&ctx, crate::StoryRange::new("body", first, first + 1))
            .unwrap();
        matches_cold("merge");
        assert_eq!(tops(&engine).len(), 2);
    }

    #[test]
    fn resident_region_fast_path_restores_moved_extents_when_it_falls_back() {
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(lowering_pages::FONT).unwrap();
        let request = small_page_request(font);
        let engine = paragraphs_engine(9607, 4);
        engine.set_local_lowering(true);
        engine.layout_document_with_regions_json(&request).unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        let lengths: Vec<u32> = engine
            .doc()
            .paragraphs("body")
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text.chars().count() as u32)
            .collect();
        let ctx = crate::EditCtx::local("", "");
        let mark = lengths[0] + 1 + lengths[1];
        engine
            .doc()
            .delete_range(&ctx, crate::StoryRange::new("body", mark, mark + 1))
            .unwrap();
        let last = lengths[0] + lengths[1] + lengths[2] + 2;
        engine
            .doc()
            .split_paragraph(&ctx, crate::Position::new("body", last + 5), None)
            .unwrap();
        assert!(
            !engine
                .apply_and_layout_regions_resident("body", &mut |_| {})
                .unwrap(),
            "a block identity change midway falls back to the full pass"
        );
        engine.apply_and_layout_regions_full().unwrap();
        assert_region_state_matches_cold(&engine, &request, "full pass after the fallback");
    }

    /// The retained region state equals a cold full pass of `request` over the same document.
    fn assert_region_state_matches_cold(engine: &EngineSession, request: &str, label: &str) {
        let snapshot = |engine: &EngineSession| {
            let pagination = engine.pagination.borrow();
            (
                serde_json::to_string(&pagination.input.as_ref().unwrap().measured).unwrap(),
                pagination.block_fingerprints.clone(),
                serde_json::to_string(&pagination.layout.as_ref().unwrap().pages).unwrap(),
            )
        };
        let resident = snapshot(engine);
        let (render, measurement, pagination, regions, display, capture, resumable) = (
            engine.render.replace(Default::default()),
            engine.measurement.replace(Default::default()),
            engine.pagination.replace(Default::default()),
            engine.regions.replace(Default::default()),
            engine.display.replace(Default::default()),
            engine.capture.replace(Default::default()),
            engine.resumable.replace(Default::default()),
        );
        engine.layout_document_with_regions_json(request).unwrap();
        assert_eq!(resident, snapshot(engine), "{label}");
        engine.render.replace(render);
        engine.measurement.replace(measurement);
        engine.pagination.replace(pagination);
        engine.regions.replace(regions);
        engine.display.replace(display);
        engine.capture.replace(capture);
        engine.resumable.replace(resumable);
    }

    #[test]
    fn appended_fonts_reuse_unaffected_blocks_and_match_a_fresh_store() {
        for (label, text, slots, bold, initial, final_chain, expected_calls) in [
            (
                "unused",
                "Latin",
                json!({}),
                false,
                json!({"requested|0|0": [0]}),
                json!({"unrelated|0|0": [1]}),
                0,
            ),
            (
                "completed chain",
                "Latin",
                json!({}),
                false,
                json!({"requested|0|0": [0]}),
                json!({"requested|0|0": [0, 1]}),
                1,
            ),
            (
                "changed chain",
                "Latin",
                json!({}),
                false,
                json!({"requested|0|0": [0]}),
                json!({"requested|0|0": [1]}),
                1,
            ),
            (
                "script slot",
                "العربية",
                json!({"cs": "Script"}),
                false,
                json!({"requested|0|0": [0]}),
                json!({"script|0|0": [1]}),
                1,
            ),
            (
                "missing alternative",
                "العربية",
                json!({"cs": "Script"}),
                true,
                json!({"requested|0|0": [0]}),
                json!({"requested|1|0": [1]}),
                1,
            ),
            (
                "missing primary",
                "Latin",
                json!({}),
                false,
                json!({}),
                json!({"requested|0|0": [1]}),
                1,
            ),
        ] {
            let fonts = docx_layout::MeasureFonts::default();
            let _scope = fonts.enter();
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                0
            );
            let engine = EngineSession::new(9651);
            let mut run = font_preflight_run(text, "Requested");
            for (slot, family) in slots.as_object().unwrap() {
                run["formatting"]["fontFamily"][slot] = family.clone();
            }
            run["formatting"]["bold"] = json!(bold);
            let blocks = [
                json!({"type": "paragraph", "content": [run]}),
                json!({"type": "paragraph", "content": [font_preflight_run("Unchanged", "Stable")]}),
            ];
            crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
            let mut request: serde_json::Value =
                serde_json::from_str(&small_page_request(0)).unwrap();
            request["measurement"]["defaults"]["fontFamily"] = json!("Stable");
            request["measurement"]["fontChains"] = initial;
            request["measurement"]["fontChains"]["stable|0|0"] = json!([0]);
            engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap();
            let mut retained = HashMap::new();
            let frame = engine.build_display_list_frame("{}", 0).unwrap();
            crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained);
            let before = engine.stats();
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                1
            );
            for (key, chain) in final_chain.as_object().unwrap() {
                request["measurement"]["fontChains"][key] = chain.clone();
            }
            engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap();
            let after = engine.stats();
            assert_eq!(
                after.resident_measure_calls - before.resident_measure_calls,
                expected_calls,
                "{label}"
            );
            assert_eq!(
                after.resident_reused_blocks - before.resident_reused_blocks,
                2 - expected_calls,
                "{label}"
            );
            let snapshot = |engine: &EngineSession| {
                let pagination = engine.pagination.borrow();
                (
                    serde_json::to_vec(&pagination.input.as_ref().unwrap().measured).unwrap(),
                    pagination.block_fingerprints.clone(),
                    serde_json::to_vec(&pagination.layout.as_ref().unwrap().pages).unwrap(),
                )
            };
            let warm = snapshot(&engine);
            let epoch = engine.display.borrow().binary_frame_epoch;
            let frame = engine.build_display_list_frame("{}", epoch).unwrap();
            let warm_pages =
                crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained).pages;
            let state = engine.doc().encode_state_as_update_v1();
            docx_layout::with_private_measure_fonts(|| {
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                    0
                );
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                    1
                );
                let cold = EngineSession::new(9652);
                cold.doc().apply_update_v1(&state).unwrap();
                cold.layout_document_with_regions_json(&request.to_string())
                    .unwrap();
                assert_eq!(warm, snapshot(&cold), "{label}");
                let frame = cold.build_display_list_frame("{}", 0).unwrap();
                let cold_pages =
                    crate::frame_delta::apply_placeholder_test_frame(&frame, &mut HashMap::new())
                        .pages;
                assert_eq!(
                    serde_json::to_vec(&warm_pages).unwrap(),
                    serde_json::to_vec(&cold_pages).unwrap(),
                    "{label}"
                );
            });
        }
    }

    #[test]
    fn prefix_font_dependencies_survive_resident_edits_and_match_a_fresh_store() {
        for missing_record in [false, true] {
            let fonts = docx_layout::MeasureFonts::default();
            let _scope = fonts.enter();
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                0
            );
            let engine = EngineSession::new(9659);
            let text = "iiiiiiiiii";
            let blocks = [
                json!({"type": "paragraph", "content": [font_preflight_run(text, "Requested")]}),
                json!({"type": "paragraph", "content": [font_preflight_run("Tail", "Stable")]}),
            ];
            crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
            let mut request: serde_json::Value =
                serde_json::from_str(&small_page_request(0)).unwrap();
            request["measurement"]["defaults"]["fontFamily"] = json!("Stable");
            request["measurement"]["fontChains"] = json!({"requested|0|0": [0], "stable|0|0": [0]});
            let prefix = engine
                .layout_document_with_regions_prefix_retained_json(&request.to_string(), 1)
                .unwrap();
            let prefix: serde_json::Value = serde_json::from_str(&prefix).unwrap();
            assert!(prefix.get("provisional").is_none());
            let chains = BTreeMap::from([
                ("requested|0|0".to_owned(), vec![1]),
                ("stable|0|0".to_owned(), vec![0]),
            ]);
            let initial_extent = {
                let pagination = engine.pagination.borrow();
                assert_eq!(pagination.input.as_ref().unwrap().measured.len(), 2);
                assert_eq!(pagination.measured_font_dependencies.len(), 2);
                assert!(
                    !pagination.measured_font_dependencies[0].matches(FontChains::BTree(&chains))
                );
                serde_json::to_vec(&pagination.input.as_ref().unwrap().measured[0].measure).unwrap()
            };
            let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
            engine.build_display_list_frame(&extras, 0).unwrap();
            if missing_record {
                engine
                    .pagination
                    .borrow_mut()
                    .measured_font_dependencies
                    .clear();
            }
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", text.len() as u32 + 1),
                    "edited ",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            let before = engine.stats();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.apply_and_layout("body", epoch).unwrap();
            let after = engine.stats();
            assert_eq!(
                after.resident_measure_calls - before.resident_measure_calls,
                1
            );
            assert_eq!(
                after.resident_reused_blocks - before.resident_reused_blocks,
                1
            );
            assert!(
                !engine.pagination.borrow().measured_font_dependencies[0]
                    .matches(FontChains::BTree(&chains))
            );

            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                1
            );
            request["measurement"]["fontChains"]["requested|0|0"] = json!([1]);
            let request_json = request.to_string();
            let before = engine.stats();
            let warm = engine
                .layout_document_with_regions_json(&request_json)
                .unwrap();
            let after = engine.stats();
            assert_eq!(
                after.resident_measure_calls - before.resident_measure_calls,
                1
            );
            assert_eq!(
                after.resident_reused_blocks - before.resident_reused_blocks,
                1
            );
            let extents = |engine: &EngineSession| {
                let pagination = engine.pagination.borrow();
                let extents: Vec<_> = pagination
                    .input
                    .as_ref()
                    .unwrap()
                    .measured
                    .iter()
                    .map(|entry| &entry.measure)
                    .collect();
                serde_json::to_vec(&extents).unwrap()
            };
            let warm_extents = extents(&engine);
            assert_ne!(
                initial_extent,
                serde_json::to_vec(
                    &engine.pagination.borrow().input.as_ref().unwrap().measured[0].measure
                )
                .unwrap()
            );
            assert!(
                engine.pagination.borrow().measured_font_dependencies[0]
                    .matches(FontChains::BTree(&chains))
            );
            let before = engine.stats();
            assert_eq!(
                engine
                    .layout_document_with_regions_json(&request_json)
                    .unwrap(),
                warm
            );
            let after = engine.stats();
            assert_eq!(after.resident_measure_calls, before.resident_measure_calls);
            assert_eq!(
                after.resident_reused_blocks - before.resident_reused_blocks,
                2
            );

            let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.build_display_list_frame(&extras, epoch).unwrap();
            let warm_pages = engine.with_display_list(|list| list.pages.clone()).unwrap();
            let state = engine.doc().encode_state_as_update_v1();
            docx_layout::with_private_measure_fonts(|| {
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                    0
                );
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                    1
                );
                let cold = EngineSession::new(9660);
                cold.doc().apply_update_v1(&state).unwrap();
                assert_eq!(
                    warm,
                    cold.layout_document_with_regions_json(&request_json)
                        .unwrap()
                );
                assert_eq!(warm_extents, extents(&cold));
                cold.build_display_list_frame(&extras, 0).unwrap();
                assert_eq!(
                    warm_pages,
                    cold.with_display_list(|list| list.pages.clone()).unwrap()
                );
            });
        }
    }

    #[test]
    fn font_store_replacement_rebuilds_display_pages_but_appends_reuse_them() {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
            0
        );
        let mut store = ooxml_text::FontStore::new();
        let font = store.register(lowering_pages::FONT.to_vec()).unwrap();
        let glyph = usize::from(store.glyph_id(font, 'z').unwrap().unwrap());
        let table_offset = |tag: &[u8; 4]| {
            let bytes = lowering_pages::FONT;
            let count = usize::from(u16::from_be_bytes(bytes[4..6].try_into().unwrap()));
            let table = bytes[12..]
                .as_chunks::<16>()
                .0
                .iter()
                .take(count)
                .find(|table| &table[..4] == tag)
                .unwrap();
            u32::from_be_bytes(table[8..12].try_into().unwrap()) as usize
        };
        let hhea = table_offset(b"hhea");
        let metric_count = usize::from(u16::from_be_bytes(
            lowering_pages::FONT[hhea + 34..hhea + 36]
                .try_into()
                .unwrap(),
        ));
        assert!(glyph < metric_count);
        let advance_offset = table_offset(b"hmtx") + glyph * 4;
        let mut replacement = lowering_pages::FONT.to_vec();
        let advance = u16::from_be_bytes(
            replacement[advance_offset..advance_offset + 2]
                .try_into()
                .unwrap(),
        );
        replacement[advance_offset..advance_offset + 2]
            .copy_from_slice(&(advance + 128).to_be_bytes());

        let engine = EngineSession::new(9661);
        let mut blocks: Vec<_> = (0..80).map(|_| json!({
            "type": "paragraph", "content": [font_preflight_run("Stable first pages", "Requested")]
        })).collect();
        blocks.push(
            json!({"type": "paragraph", "content": [font_preflight_run("zzzz", "Requested")]}),
        );
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let mut request: serde_json::Value = serde_json::from_str(&small_page_request(0)).unwrap();
        request["measurement"]["defaults"]["fontFamily"] = json!("Requested");
        request["measurement"]["fontChains"] = json!({"requested|0|0": [0]});
        let request = request.to_string();
        let extras = json!({"fontChains": {"requested|0|0": [0]}}).to_string();
        engine.layout_document_with_regions_json(&request).unwrap();
        let mut retained = HashMap::new();
        let frame = engine.build_display_list_frame(&extras, 0).unwrap();
        crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained);
        let original = engine.with_display_list(Clone::clone).unwrap();
        assert!(original.pages.len() > 2);
        assert!(
            original.pages[0]
                .primitives
                .iter()
                .any(|primitive| matches!(
                    primitive,
                    docx_layout::display_list::Primitive::GlyphRun(_)
                ))
        );
        let first_extent = serde_json::to_vec(
            &engine.pagination.borrow().input.as_ref().unwrap().measured[0].measure,
        )
        .unwrap();
        let initial_store = docx_layout::measure_store_id();
        docx_layout::clear_measure_fonts();
        assert_eq!(
            docx_layout::register_measure_font_bytes(&replacement).unwrap(),
            0
        );
        assert_ne!(docx_layout::measure_store_id(), initial_store);
        engine.layout_document_with_regions_json(&request).unwrap();
        {
            let pagination = engine.pagination.borrow();
            assert!(pagination.last_incremental);
            assert!(pagination.rebuilt_page_start > 0);
            assert_eq!(
                first_extent,
                serde_json::to_vec(&pagination.input.as_ref().unwrap().measured[0].measure)
                    .unwrap()
            );
        }
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let frame = engine.build_display_list_frame(&extras, epoch).unwrap();
        let warm_frame =
            crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained).pages;
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            after.retained_pages as u64
        );
        assert_eq!(
            original.pages[0],
            engine
                .with_display_list(|list| list.pages[0].clone())
                .unwrap()
        );

        let assert_cold = |warm_frame| {
            let state = engine.doc().encode_state_as_update_v1();
            let warm_pages = engine.with_display_list(|list| list.pages.clone()).unwrap();
            let font_count = docx_layout::measure_fonts_generation().1;
            docx_layout::with_private_measure_fonts(|| {
                assert_eq!(
                    docx_layout::register_measure_font_bytes(&replacement).unwrap(),
                    0
                );
                if font_count > 1 {
                    assert_eq!(
                        docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT)
                            .unwrap(),
                        1
                    );
                }
                let cold = EngineSession::new(9662);
                cold.doc().apply_update_v1(&state).unwrap();
                cold.layout_document_with_regions_json(&request).unwrap();
                let frame = cold.build_display_list_frame(&extras, 0).unwrap();
                let cold_frame =
                    crate::frame_delta::apply_placeholder_test_frame(&frame, &mut HashMap::new())
                        .pages;
                assert_eq!(
                    warm_pages,
                    cold.with_display_list(|list| list.pages.clone()).unwrap()
                );
                assert_eq!(
                    serde_json::to_vec(&warm_frame).unwrap(),
                    serde_json::to_vec(&cold_frame).unwrap()
                );
            });
        };
        assert_cold(warm_frame);

        let store_id = docx_layout::measure_store_id();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
            1
        );
        assert_eq!(docx_layout::measure_store_id(), store_id);
        let offset: u32 = engine
            .doc()
            .paragraphs("body")
            .unwrap()
            .iter()
            .take(80)
            .map(|paragraph| paragraph.text.len() as u32 + 1)
            .sum();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                "z",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        engine.layout_document_with_regions_json(&request).unwrap();
        assert!(engine.pagination.borrow().last_incremental);
        assert!(engine.pagination.borrow().rebuilt_page_start > 0);
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let frame = engine.build_display_list_frame(&extras, epoch).unwrap();
        let warm_frame =
            crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained).pages;
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds + 1
        );
        assert!(
            after.rebuilt_display_pages - before.rebuilt_display_pages
                < after.retained_pages as u64
        );
        assert_eq!(
            original.pages[0],
            engine
                .with_display_list(|list| list.pages[0].clone())
                .unwrap()
        );
        assert_cold(warm_frame);
    }

    #[test]
    fn referenced_font_append_rebuilds_display_pages_with_synthetic_measurements() {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
            0
        );
        let engine = EngineSession::new(9663);
        let mut blocks: Vec<_> = (0..80)
            .map(|_| {
                json!({
                    "type": "paragraph", "content": [font_preflight_run("Stable first pages", "Requested")]
                })
            })
            .collect();
        blocks.push(
            json!({"type": "paragraph", "content": [font_preflight_run("zzzz", "Requested")]}),
        );
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let mut request: serde_json::Value = serde_json::from_str(&small_page_request(0)).unwrap();
        request["measurement"]
            .as_object_mut()
            .unwrap()
            .remove("defaults");
        request["measurement"]["fontChains"] = json!({"requested|0|0": [1]});
        let request = request.to_string();
        let extras = json!({"fontChains": {"requested|0|0": [1]}}).to_string();
        engine.layout_document_with_regions_json(&request).unwrap();
        let mut retained = HashMap::new();
        let frame = engine.build_display_list_frame(&extras, 0).unwrap();
        crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained);
        let original = engine.with_display_list(Clone::clone).unwrap();
        assert!(original.pages.len() > 2);
        assert!(
            original.pages[0]
                .primitives
                .iter()
                .any(|primitive| matches!(
                    primitive,
                    docx_layout::display_list::Primitive::Text(_)
                ))
        );
        let first_extent = {
            let pagination = engine.pagination.borrow();
            let measure = &pagination.input.as_ref().unwrap().measured[0].measure;
            let BlockExtent::Paragraph(extent) = measure else {
                panic!("paragraph expected");
            };
            assert_eq!(extent.lines[0].synthetic_fallback, Some(true));
            serde_json::to_vec(measure).unwrap()
        };
        let store_id = docx_layout::measure_store_id();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
            1
        );
        assert_eq!(docx_layout::measure_store_id(), store_id);
        let offset: u32 = engine
            .doc()
            .paragraphs("body")
            .unwrap()
            .iter()
            .take(80)
            .map(|paragraph| paragraph.text.len() as u32 + 1)
            .sum();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                "z",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let warm_layout = engine.layout_document_with_regions_json(&request).unwrap();
        {
            let pagination = engine.pagination.borrow();
            assert!(pagination.last_incremental);
            assert!(pagination.rebuilt_page_start > 0);
            assert_eq!(
                first_extent,
                serde_json::to_vec(&pagination.input.as_ref().unwrap().measured[0].measure)
                    .unwrap()
            );
        }
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let frame = engine.build_display_list_frame(&extras, epoch).unwrap();
        let warm_frame =
            crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained).pages;
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            after.retained_pages as u64
        );
        let warm_pages = engine.with_display_list(|list| list.pages.clone()).unwrap();
        assert_ne!(original.pages[0], warm_pages[0]);
        assert!(warm_pages[0].primitives.iter().any(|primitive| matches!(
            primitive,
            docx_layout::display_list::Primitive::GlyphRun(_)
        )));
        let state = engine.doc().encode_state_as_update_v1();
        docx_layout::with_private_measure_fonts(|| {
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                0
            );
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                1
            );
            let cold = EngineSession::new(9664);
            cold.doc().apply_update_v1(&state).unwrap();
            assert_eq!(
                warm_layout,
                cold.layout_document_with_regions_json(&request).unwrap()
            );
            let frame = cold.build_display_list_frame(&extras, 0).unwrap();
            let cold_frame =
                crate::frame_delta::apply_placeholder_test_frame(&frame, &mut HashMap::new()).pages;
            assert_eq!(
                warm_pages,
                cold.with_display_list(|list| list.pages.clone()).unwrap()
            );
            assert_eq!(
                serde_json::to_vec(&warm_frame).unwrap(),
                serde_json::to_vec(&cold_frame).unwrap()
            );
        });
    }

    #[test]
    fn first_font_invalidates_resident_measurements_without_a_matching_chain() {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        let engine = EngineSession::new(9653);
        let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
            "kind": "paragraph", "id": "empty-run", "runs": [
                {"kind": "text", "text": "", "fontFamily": "Requested"},
                {"kind": "lineBreak"}
            ]
        }]))
        .unwrap();
        let mut request = json!({
            "measurement": {"defaults": {"fontFamily": "Requested", "fontSize": 12}}
        });
        let prepared = engine
            .prepare_region_layout(&request.to_string(), None)
            .unwrap();
        let (empty, dependencies) = FontChainDependencies::capture(|| {
            docx_layout::measure_blocks::measure_block(&mut blocks[0], 300.0, &prepared.measurement)
                .unwrap()
        });
        {
            let mut pagination = engine.pagination.borrow_mut();
            pagination.input = Some(LayoutInput {
                measured: vec![MeasuredBlock {
                    block: blocks[0].clone(),
                    measure: empty.clone(),
                }],
                options: Default::default(),
            });
            pagination.block_fingerprints =
                measured_fingerprints(pagination.input.as_ref().unwrap()).unwrap();
            pagination.measured_with = Some(prepared.measurement_fingerprint);
            pagination.measured_widths = vec![300.0];
            pagination.measured_font_dependencies = vec![dependencies.clone()];
        }
        let (reused, fingerprints) = engine
            .resident_region_measured(
                &blocks,
                &[300.0],
                &[],
                &prepared.regions,
                &prepared.measurement,
                prepared.measurement_fingerprint,
                None,
            )
            .unwrap()
            .unwrap();
        {
            let mut pagination = engine.pagination.borrow_mut();
            pagination.input.as_mut().unwrap().measured = reused;
            pagination.block_fingerprints = fingerprints;
        }
        let font = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        assert_eq!(docx_layout::register_measure_font_bytes(font).unwrap(), 0);
        request["measurement"]["fontChains"] = json!({"unrelated|0|0": [0]});
        let prepared = engine
            .prepare_region_layout(&request.to_string(), None)
            .unwrap();
        assert!(dependencies.matches(FontChains::BTree(&prepared.measurement.font_chains)));
        assert!(
            engine
                .resident_region_measured(
                    &blocks,
                    &[300.0],
                    &[],
                    &prepared.regions,
                    &prepared.measurement,
                    prepared.measurement_fingerprint,
                    None,
                )
                .unwrap()
                .is_none()
        );
        let warm = docx_layout::measure_blocks::measure_block(
            &mut blocks[0],
            300.0,
            &prepared.measurement,
        )
        .unwrap();
        assert_ne!(
            serde_json::to_vec(&empty).unwrap(),
            serde_json::to_vec(&warm).unwrap()
        );
        let cold = docx_layout::with_private_measure_fonts(|| {
            assert_eq!(docx_layout::register_measure_font_bytes(font).unwrap(), 0);
            docx_layout::measure_blocks::measure_block(&mut blocks[0], 300.0, &prepared.measurement)
                .unwrap()
        });
        assert_eq!(
            serde_json::to_vec(&warm).unwrap(),
            serde_json::to_vec(&cold).unwrap()
        );
    }

    #[test]
    fn registered_chain_font_invalidates_synthetic_resident_measurements() {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
            0
        );
        let engine = EngineSession::new(9657);
        let blocks = [json!({
            "type": "paragraph", "content": [font_preflight_run("Latin", "Requested")]
        })];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let mut request: serde_json::Value = serde_json::from_str(&small_page_request(0)).unwrap();
        request["measurement"]["defaults"]["fontFamily"] = json!("Requested");
        request["measurement"]["fontChains"] = json!({"requested|0|0": [1]});
        let request = request.to_string();
        engine.layout_document_with_regions_json(&request).unwrap();
        let snapshot = |engine: &EngineSession| {
            let pagination = engine.pagination.borrow();
            (
                serde_json::to_vec(&pagination.input.as_ref().unwrap().measured).unwrap(),
                pagination.block_fingerprints.clone(),
                serde_json::to_vec(&pagination.layout.as_ref().unwrap().pages).unwrap(),
            )
        };
        let synthetic = snapshot(&engine);
        let dependencies = {
            let pagination = engine.pagination.borrow();
            let BlockExtent::Paragraph(extent) =
                &pagination.input.as_ref().unwrap().measured[0].measure
            else {
                panic!("paragraph expected");
            };
            assert_eq!(extent.lines[0].synthetic_fallback, Some(true));
            pagination.measured_font_dependencies[0].clone()
        };
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
            1
        );
        let chains = BTreeMap::from([("requested|0|0".to_owned(), vec![1])]);
        assert!(dependencies.matches(FontChains::BTree(&chains)));
        engine.layout_document_with_regions_json(&request).unwrap();
        let warm = snapshot(&engine);
        assert_ne!(synthetic.0, warm.0);
        {
            let pagination = engine.pagination.borrow();
            let BlockExtent::Paragraph(extent) =
                &pagination.input.as_ref().unwrap().measured[0].measure
            else {
                panic!("paragraph expected");
            };
            assert_ne!(extent.lines[0].synthetic_fallback, Some(true));
        }
        let state = engine.doc().encode_state_as_update_v1();
        docx_layout::with_private_measure_fonts(|| {
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                0
            );
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                1
            );
            let cold = EngineSession::new(9658);
            cold.doc().apply_update_v1(&state).unwrap();
            cold.layout_document_with_regions_json(&request).unwrap();
            assert_eq!(warm, snapshot(&cold));
        });
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
            2
        );
        let before = engine.stats();
        engine.layout_document_with_regions_json(&request).unwrap();
        let after = engine.stats();
        assert_eq!(after.resident_measure_calls, before.resident_measure_calls);
        assert_eq!(
            after.resident_reused_blocks - before.resident_reused_blocks,
            1
        );
        assert_eq!(warm, snapshot(&engine));
    }

    #[test]
    fn appended_fonts_keep_the_line_and_dependencies_of_an_empty_section() {
        for with_previous_section in [false, true] {
            let fonts = docx_layout::MeasureFonts::default();
            let _scope = fonts.enter();
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                0
            );
            let engine = EngineSession::new(9654);
            let mut blocks = Vec::new();
            if with_previous_section {
                blocks.push(json!({
                    "type": "paragraph", "content": [font_preflight_run("Prefix", "Stable")],
                    "sectionProperties": {"sectionStart": "nextPage"}
                }));
            }
            blocks.extend([
                json!({
                    "type": "paragraph", "content": [],
                    "formatting": {"runProperties": {"fontFamily": {"ascii": "Requested"}}},
                    "sectionProperties": {"sectionStart": "nextPage"}
                }),
                json!({"type": "paragraph", "content": [font_preflight_run("Tail", "Stable")]}),
            ]);
            crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
            let mut request: serde_json::Value =
                serde_json::from_str(&small_page_request(0)).unwrap();
            request["measurement"]["defaults"]["fontFamily"] = json!("Stable");
            request["measurement"]["fontChains"] = json!({"requested|0|0": [0], "stable|0|0": [0]});
            let section = request["regions"]["sections"][0].clone();
            for _ in 0..1 + usize::from(with_previous_section) {
                request["regions"]["sections"]
                    .as_array_mut()
                    .unwrap()
                    .push(section.clone());
            }
            engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap();
            let opening = if with_previous_section { 2 } else { 0 };
            let assert_line = |engine: &EngineSession| {
                let pagination = engine.pagination.borrow();
                let BlockExtent::Paragraph(extent) =
                    &pagination.input.as_ref().unwrap().measured[opening].measure
                else {
                    panic!("empty section must retain a paragraph");
                };
                assert_eq!(extent.lines.len(), 1);
                assert!(extent.total_height > 0.0);
            };
            assert_line(&engine);
            let mut retained = HashMap::new();
            let frame = engine.build_display_list_frame("{}", 0).unwrap();
            crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained);
            let before = engine.stats();
            assert_eq!(
                docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                1
            );
            request["measurement"]["fontChains"]["requested|0|0"] = json!([1]);
            let warm = engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap();
            assert_eq!(
                engine.stats().resident_measure_calls - before.resident_measure_calls,
                1
            );
            assert_line(&engine);
            let dependencies =
                engine.pagination.borrow().measured_font_dependencies[opening].clone();
            let mut measurement: docx_layout::measure_blocks::MeasurementConfig =
                serde_json::from_value(request["measurement"].clone()).unwrap();
            assert!(dependencies.matches(FontChains::BTree(&measurement.font_chains)));
            measurement.font_chains.remove("requested|0|0");
            assert!(!dependencies.matches(FontChains::BTree(&measurement.font_chains)));
            let epoch = engine.display.borrow().binary_frame_epoch;
            let frame = engine.build_display_list_frame("{}", epoch).unwrap();
            let warm_pages =
                crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained).pages;
            let state = engine.doc().encode_state_as_update_v1();
            docx_layout::with_private_measure_fonts(|| {
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
                    0
                );
                assert_eq!(
                    docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
                    1
                );
                let cold = EngineSession::new(9655);
                cold.doc().apply_update_v1(&state).unwrap();
                assert_eq!(
                    warm,
                    cold.layout_document_with_regions_json(&request.to_string())
                        .unwrap()
                );
                assert_line(&cold);
                let frame = cold.build_display_list_frame("{}", 0).unwrap();
                let cold_pages =
                    crate::frame_delta::apply_placeholder_test_frame(&frame, &mut HashMap::new())
                        .pages;
                assert_eq!(
                    serde_json::to_vec(&warm_pages).unwrap(),
                    serde_json::to_vec(&cold_pages).unwrap()
                );
            });
        }
    }

    #[test]
    fn resident_section_marks_remeasure_when_the_previous_block_kind_changes() {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap(),
            0
        );
        for previous_is_section in [false, true] {
            let engine = EngineSession::new(9656);
            let prepared = engine
                .prepare_region_layout(
                    &json!({
                        "measurement": {
                            "fontChains": {"requested|0|0": [0]},
                            "defaults": {"fontFamily": "Requested", "fontSize": 12}
                        }
                    })
                    .to_string(),
                    None,
                )
                .unwrap();
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind": if previous_is_section {"sectionBreak"} else {"pageBreak"}, "id": "previous"},
                {"kind": "paragraph", "id": "mark", "runs": []},
                {"kind": "sectionBreak", "id": "next"}
            ])).unwrap();
            let widths = [300.0; 3];
            let mut flow = docx_layout::measure_blocks::FloatFlow::new(
                &blocks,
                &widths,
                &prepared.measurement,
                None,
            )
            .unwrap();
            flow.measure_until(&mut blocks, &widths, &prepared.measurement, 3)
                .unwrap();
            {
                let mut pagination = engine.pagination.borrow_mut();
                pagination.measured_font_dependencies = flow.font_dependencies().to_vec();
                pagination.input = Some(LayoutInput {
                    measured: blocks
                        .iter()
                        .cloned()
                        .zip(flow.into_extents())
                        .map(|(block, measure)| MeasuredBlock { block, measure })
                        .collect(),
                    options: Default::default(),
                });
                pagination.block_fingerprints =
                    measured_fingerprints(pagination.input.as_ref().unwrap()).unwrap();
                pagination.measured_with = Some(prepared.measurement_fingerprint);
                pagination.measured_widths = widths.to_vec();
            }
            blocks[0] = serde_json::from_value(json!({
                "kind": if previous_is_section {"pageBreak"} else {"sectionBreak"}, "id": "previous"
            }))
            .unwrap();
            let (warm, _) = engine
                .resident_region_measured(
                    &blocks,
                    &widths,
                    &[],
                    &prepared.regions,
                    &prepared.measurement,
                    prepared.measurement_fingerprint,
                    None,
                )
                .unwrap()
                .unwrap();
            assert_eq!(engine.stats().resident_measure_calls, 2);
            let BlockExtent::Paragraph(extent) = &warm[1].measure else {
                panic!("section mark must retain a paragraph");
            };
            assert_eq!(extent.lines.len(), usize::from(!previous_is_section));
            let dependencies = engine.pagination.borrow().measured_font_dependencies[1].clone();
            assert_eq!(
                dependencies.matches(FontChains::BTree(&BTreeMap::new())),
                previous_is_section
            );
            let mut cold = blocks.clone();
            let extents = docx_layout::measure_blocks::measure_blocks_with_floats(
                &mut cold,
                &widths,
                &prepared.measurement,
                None,
            )
            .unwrap();
            let warm_extents: Vec<_> = warm.iter().map(|entry| &entry.measure).collect();
            assert_eq!(
                serde_json::to_vec(&warm_extents).unwrap(),
                serde_json::to_vec(&extents).unwrap()
            );
        }
    }

    fn paragraphs_engine(client_id: u64, paragraphs: usize) -> EngineSession {
        let engine = EngineSession::new(client_id);
        engine
            .doc()
            .create_story("body", "", "Normal", "left")
            .unwrap();
        let ctx = crate::EditCtx::local("", "");
        let mut cursor = 0_u32;
        for index in 0..paragraphs {
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
            if index + 1 < paragraphs {
                engine
                    .doc()
                    .split_paragraph(&ctx, crate::Position::new("body", cursor), None)
                    .unwrap();
                cursor += 1;
            }
        }
        engine
    }

    /// Appends the final section again, as hosts build region requests.
    fn repeat_final_section(request: &mut serde_json::Value) {
        let sections = request["regions"]["sections"].as_array_mut().unwrap();
        sections.push(sections.last().unwrap().clone());
    }

    fn small_page_request(font_id: u32) -> String {
        serde_json::json!({
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
        })
        .to_string()
    }

    fn local_patch_mixed_runs() -> String {
        use super::lowering_fixture::run;
        format!(
            r#"{}<w:r><w:rPr><w:b/></w:rPr><w:t>cd</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ef</w:t></w:r>{}"#,
            run("ab"),
            run("gh")
        )
    }

    fn local_patch_list_paragraph(
        id: &str,
        num: u32,
        level: u32,
        mark: &str,
        content: &str,
    ) -> String {
        format!(
            r#"<w:p w14:paraId="{id}"><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num}"/></w:numPr>{mark}</w:pPr>{content}</w:p>"#
        )
    }

    fn local_patch_numbering(format: &str, marker: &str) -> String {
        format!(
            r#"<w:numbering {}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{format}"/><w:lvlText w:val="{marker}"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>"#,
            lowering_fixture::NS
        )
    }

    fn local_patch_laid_out(
        bytes: &[u8],
        client_id: u64,
        enabled: bool,
    ) -> (EngineSession, String) {
        let (engine, request) = lowering_pages::laid_out(bytes, client_id);
        engine.set_local_lowering(enabled);
        engine.render.replace(Default::default());
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        (engine, request)
    }

    fn local_patch_step(
        engine: &EngineSession,
        request: &str,
        story: &str,
        (start, end, text): (u32, u32, Option<&str>),
        patched: bool,
    ) {
        use crate::{Position, StoryRange};

        let enabled = engine.local_lowering.get();
        let snapshot = |engine: &EngineSession| {
            let render = engine.render.borrow();
            let lowered = &render.stories["body"];
            let pagination = engine.pagination.borrow();
            (
                serde_json::to_string(lowered.blocks.as_ref()).unwrap(),
                lowered.map.as_ref().clone(),
                serde_json::to_string(&pagination.input.as_ref().unwrap().measured).unwrap(),
                pagination.block_fingerprints.clone(),
                serde_json::to_string(&pagination.layout.as_ref().unwrap().pages).unwrap(),
            )
        };
        let before = Rc::as_ptr(&engine.render.borrow().stories["body"].blocks);
        if text.is_none() && start == end {
            let ctx = crate::EditCtx::local("", "");
            let position = Position::new(story, start);
            engine.doc().split_paragraph(&ctx, position, None).unwrap();
        } else {
            let range = StoryRange::new(story, start, end);
            engine.edit_resident_text(range, text, true).unwrap();
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine
            .apply_and_layout(story, epoch)
            .unwrap_or_else(|error| panic!("{story} [{start}, {end}) {text:?}: {error}"));
        let after = Rc::as_ptr(&engine.render.borrow().stories["body"].blocks);
        let incremental = snapshot(engine);
        macro_rules! cold {
            ($($field:ident)+) => {{
                let ($($field,)+) = ($(engine.$field.replace(Default::default()),)+);
                engine.layout_document_with_regions_json(request).unwrap();
                let result = snapshot(engine);
                $(engine.$field.replace($field);)+
                result
            }};
        }
        let oracle = cold!(render measurement pagination regions display capture resumable);
        assert_eq!(incremental, oracle, "{story} [{start}, {end}) {text:?}");
        assert_eq!(
            before == after,
            patched && enabled,
            "{story} [{start}, {end}) {text:?} enabled={enabled}"
        );
    }

    #[test]
    fn resident_plain_text_patch_matches_cold_full() {
        for enabled in [false, true] {
            resident_plain_text_patch_matches_cold_full_in(enabled);
        }
    }

    #[test]
    fn resident_text_insert_windowed_frames_match_with_local_lowering() {
        use yrs::{Assoc, IndexedSequence};

        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(lowering_pages::FONT).unwrap();
        let request = small_page_request(font);
        let [(full, full_frame), (local, local_frame)] = [false, true].map(|enabled| {
            let engine = paragraphs_engine(9608, 24);
            engine.set_local_lowering(enabled);
            engine.layout_document_with_regions_json(&request).unwrap();
            let paragraph = engine.doc().paragraphs("body").unwrap().remove(12);
            let paragraphs = engine.doc().paragraph_index("body").unwrap();
            let (start, _) = paragraphs.para_span(&paragraph.para_id).unwrap();
            let txn = engine.doc().yrs_doc().transact();
            let story = crate::story_ref(&txn, "body").unwrap();
            let head = story.sticky_index(&txn, start + 13, Assoc::After).unwrap();
            drop(txn);
            engine.set_resident_caret_head(Some(("body".to_owned(), head)));
            engine.set_display_window(Some(0..1));
            engine.set_windowed_incremental_builds(true);
            let frame = engine.build_display_list_frame("{}", 0).unwrap();
            (engine, frame)
        });
        assert_eq!(local_frame, full_frame);
        for text in ["x", "😀", "y"] {
            let frames = [&full, &local].map(|engine| {
                let txn = engine.doc().yrs_doc().transact();
                let index = engine
                    .resident_caret_head
                    .borrow()
                    .as_ref()
                    .unwrap()
                    .1
                    .get_offset(&txn)
                    .unwrap()
                    .index;
                drop(txn);
                engine
                    .edit_resident_text(
                        crate::StoryRange::new("body", index, index),
                        Some(text),
                        true,
                    )
                    .unwrap();
                let epoch = engine.display.borrow().binary_frame_epoch;
                engine.apply_and_layout("body", epoch).unwrap()
            });
            assert_eq!(frames[1], frames[0], "insert {text:?}");
            assert!(
                local
                    .with_display_list(|list| list.pages.iter().any(|page| page.unbuilt))
                    .unwrap()
            );
        }
        docx_layout::clear_measure_fonts();
    }

    fn resident_plain_text_patch_matches_cold_full_in(enabled: bool) {
        use super::lowering_fixture::{Package, para, run};
        let laid_out = |bytes: &[u8], client_id| local_patch_laid_out(bytes, client_id, enabled);
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(lowering_pages::FONT).unwrap();
        let mut repeated: serde_json::Value =
            serde_json::from_str(&small_page_request(font)).unwrap();
        repeat_final_section(&mut repeated);
        let step = local_patch_step;
        for request in [small_page_request(font), repeated.to_string()] {
            let engine = paragraphs_engine(9600, 3);
            engine.set_local_lowering(enabled);
            engine.layout_document_with_regions_json(&request).unwrap();
            engine.build_display_list_frame("{}", 0).unwrap();
            for (paragraph, offset) in (0..3).flat_map(|p| [0, 13, u32::MAX].map(|at| (p, at))) {
                let paragraphs = engine.doc().paragraphs("body").unwrap();
                let width = paragraphs[0].text.encode_utf16().count() as u32;
                let start = paragraph as u32 * (width + 1);
                let at = start + offset.min(width);
                step(&engine, &request, "body", (at, at, Some("😀")), true);
                step(&engine, &request, "body", (at, at + 2, None), true);
                step(&engine, &request, "body", (at, at, Some("x")), true);
                step(&engine, &request, "body", (at, at + 1, None), true);
            }
            let wrapping = " wrap".repeat(250);
            step(&engine, &request, "body", (2, 2, Some(&wrapping)), true);
            step(&engine, &request, "body", (2, 4, None), false);
            step(&engine, &request, "body", (2, 2, None), false);
        }
        let request = repeated.to_string();
        let empty = EngineSession::new(9601);
        empty.set_local_lowering(enabled);
        let doc = empty.doc();
        doc.create_story("body", "A", "Normal", "left").unwrap();
        empty.layout_document_with_regions_json(&request).unwrap();
        empty.build_display_list_frame("{}", 0).unwrap();
        for (end, text) in [(1, None), (0, Some("😀")), (2, None), (0, Some("B"))] {
            step(&empty, &request, "body", (0, end, text), true);
        }
        let cell = |id, text| format!("<w:tc><w:tcPr/>{}</w:tc>", para(id, &run(text)));
        let table_body = format!(
            r#"{}<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>{}"#,
            para("10000001", &run("Before")),
            cell("10000002", "First cell"),
            cell("10000003", "Second cell"),
            para("10000004", &run("After"))
        );
        let numbered = r#"<w:p w14:paraId="10000001"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List</w:t></w:r></w:p>"#;
        let numbering = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#;
        let packages = [
            ("table", Package::new(&table_body)),
            ("list", Package::new(numbered).numbering(numbering)),
            ("field", Package::new(&para("10000001", r#"<w:fldSimple w:instr=" SEQ Example "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#))),
            ("bookmark", Package::new(&para("10000001", r#"<w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t>Book</w:t></w:r><w:bookmarkEnd w:id="0"/>"#))),
            ("comment", Package::new(&para("10000001", r#"<w:commentRangeStart w:id="0"/><w:r><w:t>Comment</w:t></w:r><w:commentRangeEnd w:id="0"/>"#)).part(
                "comments.xml", "rIdComments", "comments", "comments",
                &format!(r#"<w:comments {}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"#, lowering_fixture::NS))),
            ("tracked", Package::new(&para("10000001", r#"<w:ins w:id="1" w:author="A"><w:r><w:t>Change</w:t></w:r></w:ins>"#))),
            ("formatting", Package::new(&para("10000001", r#"<w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>Changed</w:t></w:r>"#))),
            ("mixed", Package::new(&para("10000001", r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t>Plain</w:t></w:r>"#))),
            ("sections", Package::new(&format!(r#"<w:p w14:paraId="10000001"><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>{}"#, para("10000002", &run("Second"))))),
            ("contextual", Package::new(&format!(r#"<w:p w14:paraId="10000001"><w:pPr><w:contextualSpacing/></w:pPr><w:r><w:t>Before</w:t></w:r></w:p>{}"#, para("10000002", &run("After"))))),
        ];
        for (name, package) in packages {
            let (engine, request) = laid_out(&package.bytes(), 9602);
            let patched = matches!(name, "table" | "list" | "mixed");
            step(&engine, &request, "body", (0, 0, Some("x")), patched);
            if name == "table" {
                step(&engine, &request, "body", (11, 11, Some("😀")), true);
                step(&engine, &request, "body:t0:r0c1", (2, 2, Some("x")), false);
            } else if name == "contextual" {
                step(&engine, &request, "body", (8, 8, Some("x")), false);
            }
        }
        let bold = Package::new(&para(
            "10000001",
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>"#,
        ));
        let (engine, request) = laid_out(&bold.bytes(), 9604);
        step(&engine, &request, "body", (2, 2, Some("x")), true);
        step(&engine, &request, "body", (0, 1, None), true);
        let between = |content: &str| {
            Package::new(&format!(
                "{}{}{}",
                para("10000001", &run("Before")),
                para("10000002", content),
                para("10000003", &run("After"))
            ))
        };
        let list = |middle: &str, mark: &str, last_num| {
            format!(
                "{}{}{}",
                local_patch_list_paragraph("10000001", 1, 0, "", &run("First")),
                local_patch_list_paragraph("10000002", 1, 0, mark, middle),
                local_patch_list_paragraph("10000003", last_num, 0, "", &run("Last"))
            )
        };
        let decimal = local_patch_numbering("decimal", "%1.");
        let bullet = local_patch_numbering("bullet", "•");
        let multilevel: String = [
            "First", "Nested", "Second", "Middle", "Third", "Nested", "Last",
        ]
        .iter()
        .enumerate()
        .map(|(index, text)| {
            local_patch_list_paragraph(
                &format!("{:08X}", 0x10000001 + index),
                1,
                (index % 2) as u32,
                "",
                &run(text),
            )
        })
        .collect();
        let cases: Vec<(&str, Package, usize, &[u32])> = vec![
            ("mixed", between(&local_patch_mixed_runs()), 1, &[]),
            (
                "rsid",
                between(
                    r#"<w:r w:rsidR="00000001"><w:t>ab</w:t></w:r><w:r w:rsidR="00000002"><w:t>cd</w:t></w:r>"#,
                ),
                1,
                &[],
            ),
            (
                "language",
                between(
                    r#"<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:lang w:val="fr-FR"/></w:rPr><w:t>cd</w:t></w:r>"#,
                ),
                1,
                &[],
            ),
            (
                "proofing",
                between(&format!(
                    r#"{}<w:proofErr w:type="spellStart"/>{}<w:proofErr w:type="spellEnd"/>"#,
                    run("ab"),
                    run("cd")
                )),
                1,
                &[],
            ),
            (
                "one-character segment",
                between(&format!(
                    r#"{}<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>{}"#,
                    run("ab"),
                    run("cd")
                )),
                1,
                &[2],
            ),
            (
                "bullet",
                Package::new(&list(&run("Middle"), "", 1)).numbering(&bullet),
                1,
                &[],
            ),
            (
                "restart",
                Package::new(&list(&run("Middle"), "", 2)).numbering(&decimal),
                1,
                &[],
            ),
            (
                "multilevel",
                Package::new(&multilevel).numbering(&decimal),
                3,
                &[],
            ),
            (
                "marker formatting",
                Package::new(&list(
                    &run("Middle"),
                    r#"<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>"#,
                    1,
                ))
                .numbering(&decimal),
                1,
                &[],
            ),
            (
                "tracked insertion",
                between(&format!(
                    r#"{}<w:ins w:id="1" w:author="A">{}</w:ins>{}"#,
                    run("ab"),
                    run("cd"),
                    run("ef")
                )),
                1,
                &[],
            ),
            (
                "tracked deletion",
                between(&format!(
                    r#"{}<w:del w:id="1" w:author="A"><w:r><w:delText>cd</w:delText></w:r></w:del>{}"#,
                    run("ab"),
                    run("ef")
                )),
                1,
                &[],
            ),
            (
                "empty list",
                Package::new(&list("", "", 1)).numbering(&decimal),
                1,
                &[],
            ),
        ];
        for (name, package, paragraph, fallback_deletes) in cases {
            let bytes = package.bytes();
            let (initial, _) = laid_out(&bytes, 9607);
            let paragraphs = initial.doc().paragraphs("body").unwrap();
            let start: u32 = paragraphs[..paragraph]
                .iter()
                .map(|paragraph| paragraph.text.encode_utf16().count() as u32 + 1)
                .sum();
            let text = &paragraphs[paragraph].text;
            let width = text.encode_utf16().count() as u32;
            assert!(
                text.is_ascii(),
                "{name}: every UTF-16 offset is a scalar boundary"
            );
            if name == "marker formatting" {
                let render = initial.render.borrow();
                let LayoutBlock::Paragraph(block) = &render.stories["body"].blocks[paragraph]
                else {
                    panic!("list paragraph expected");
                };
                let attrs = block.attrs.as_ref().unwrap();
                assert_eq!(attrs.list_marker.as_deref(), Some("2."));
                assert_eq!(attrs.list_marker_bold, Some(true));
                assert_eq!(attrs.list_marker_color.as_deref(), Some("#FF0000"));
            }
            let patched = !name.starts_with("tracked");
            for offset in 0..=width {
                for inserted in ["x", "😀"] {
                    let (engine, request) = laid_out(&bytes, 9608);
                    let at = start + offset;
                    step(&engine, &request, "body", (at, at, Some(inserted)), patched);
                    step(
                        &engine,
                        &request,
                        "body",
                        (at, at + inserted.encode_utf16().count() as u32, None),
                        patched,
                    );
                }
            }
            for offset in 0..width {
                let (engine, request) = laid_out(&bytes, 9609);
                let at = start + offset;
                step(
                    &engine,
                    &request,
                    "body",
                    (at, at + 1, None),
                    patched && !fallback_deletes.contains(&offset),
                );
            }
        }
        for (bytes, patched) in [
            (
                include_bytes!(
                    "../tests/fixtures/field-code-paragraphs/body-field-code-paragraphs.docx"
                )
                .as_slice(),
                false,
            ),
            (
                include_bytes!("../tests/fixtures/suppressed-list-markers.docx").as_slice(),
                true,
            ),
            (
                include_bytes!("../tests/fixtures/page-fragments/pages.docx").as_slice(),
                false,
            ),
            (
                include_bytes!("../tests/fixtures/footnote-anchor.docx").as_slice(),
                false,
            ),
        ] {
            let (engine, request) = laid_out(bytes, 9603);
            step(&engine, &request, "body", (0, 0, Some("x")), patched);
        }
    }

    /// Earlier inserts shift mixed and list seeds without forcing full lowering.
    #[test]
    fn resident_shifted_mixed_and_list_text_patches_match_cold_full() {
        use super::lowering_fixture::{Package, para, run};

        let step = local_patch_step;
        let body = format!(
            "{}{}{}",
            para("10000001", &run("Before")),
            para("10000002", &local_patch_mixed_runs()),
            local_patch_list_paragraph("10000003", 1, 0, "", &local_patch_mixed_runs())
        );
        for (format, marker) in [("bullet", "•"), ("decimal", "%1.")] {
            let bytes = Package::new(&body)
                .numbering(&local_patch_numbering(format, marker))
                .bytes();
            let (engine, request) = local_patch_laid_out(&bytes, 9611, true);
            step(&engine, &request, "body", (2, 2, Some("😀x")), true);
            for paragraph in [1, 2] {
                for offset in [0, 4, u32::MAX] {
                    let paragraphs = engine.doc().paragraphs("body").unwrap();
                    let start: u32 = paragraphs[..paragraph]
                        .iter()
                        .map(|paragraph| paragraph.text.encode_utf16().count() as u32 + 1)
                        .sum();
                    let width = paragraphs[paragraph].text.encode_utf16().count() as u32;
                    let at = start + offset.min(width);
                    for inserted in ["x", "😀"] {
                        step(&engine, &request, "body", (at, at, Some(inserted)), true);
                        step(
                            &engine,
                            &request,
                            "body",
                            (at, at + inserted.encode_utf16().count() as u32, None),
                            true,
                        );
                    }
                }
            }
        }
    }

    /// Supplementary characters at formatting boundaries obey scalar patch limits.
    #[test]
    fn resident_mixed_surrogate_boundary_text_patches_match_cold_full() {
        use super::lowering_fixture::{Package, para, run};

        let step = local_patch_step;
        let bytes = Package::new(&para(
            "10000001",
            &format!(
                r#"{}<w:r><w:rPr><w:b/></w:rPr><w:t>😀</w:t></w:r>{}"#,
                run("ab"),
                run("😀cd")
            ),
        ))
        .bytes();
        let (initial, _) = local_patch_laid_out(&bytes, 9612, true);
        let paragraphs = initial.doc().paragraphs("body").unwrap();
        assert_eq!(paragraphs[0].text, "ab😀😀cd");
        let width = paragraphs[0].text.encode_utf16().count() as u32;
        for offset in 0..=width {
            for inserted in ["x", "😀"] {
                let (engine, request) = local_patch_laid_out(&bytes, 9612, true);
                let patched = !matches!(offset, 3 | 5);
                step(
                    &engine,
                    &request,
                    "body",
                    (offset, offset, Some(inserted)),
                    patched,
                );
                if patched {
                    step(
                        &engine,
                        &request,
                        "body",
                        (
                            offset,
                            offset + inserted.encode_utf16().count() as u32,
                            None,
                        ),
                        true,
                    );
                }
            }
            for removed in [1, 2] {
                if offset + removed > width {
                    continue;
                }
                let (engine, request) = local_patch_laid_out(&bytes, 9612, true);
                let patched = matches!(
                    (offset, removed),
                    (0, 1) | (1, 1) | (4, 2) | (6, 1) | (7, 1)
                );
                step(
                    &engine,
                    &request,
                    "body",
                    (offset, offset + removed, None),
                    patched,
                );
            }
        }
        for (start, end) in [(0, 2), (1, 4), (2, 6), (4, 8)] {
            let (engine, request) = local_patch_laid_out(&bytes, 9612, true);
            step(&engine, &request, "body", (start, end, None), false);
        }
    }

    #[test]
    fn resident_mixed_and_list_text_patch_frames_match_full_lowering() {
        use super::lowering_fixture::{Package, para, run};
        use crate::StoryRange;

        let body = format!(
            "{}{}{}{}",
            para("10000001", &local_patch_mixed_runs()),
            local_patch_list_paragraph("10000002", 1, 0, "", &run("First")),
            local_patch_list_paragraph(
                "10000003",
                1,
                0,
                r#"<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>"#,
                &run("Middle"),
            ),
            local_patch_list_paragraph("10000004", 2, 0, "", &run("Last"))
        );
        let bytes = Package::new(&body)
            .numbering(&local_patch_numbering("decimal", "%1."))
            .bytes();
        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(lowering_pages::FONT).unwrap();
        let request = small_page_request(font);
        let engines = [EngineSession::new(9610), EngineSession::new(9610)];
        for (index, engine) in engines.iter().enumerate() {
            crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
            engine.set_local_lowering(index == 1);
            engine.layout_document_with_regions_json(&request).unwrap();
        }
        let frames = || {
            engines
                .iter()
                .map(|engine| engine.build_display_list_frame("{}", 0).unwrap())
                .collect::<Vec<_>>()
        };
        let initial = frames();
        assert_eq!(initial[0], initial[1]);
        for paragraph in [0, 2] {
            let width = engines[0].doc().paragraphs("body").unwrap()[paragraph]
                .text
                .encode_utf16()
                .count() as u32;
            for offset in 0..=width {
                for inserted in ["x", "😀"] {
                    let paragraphs = engines[0].doc().paragraphs("body").unwrap();
                    let start: u32 = paragraphs[..paragraph]
                        .iter()
                        .map(|paragraph| paragraph.text.encode_utf16().count() as u32 + 1)
                        .sum();
                    let at = start + offset;
                    for (end, text) in [
                        (at, Some(inserted)),
                        (at + inserted.encode_utf16().count() as u32, None),
                    ] {
                        let before = Rc::as_ptr(&engines[1].render.borrow().stories["body"].blocks);
                        for engine in &engines {
                            engine
                                .edit_resident_text(StoryRange::new("body", at, end), text, true)
                                .unwrap();
                            let epoch = engine.display.borrow().binary_frame_epoch;
                            engine.apply_and_layout("body", epoch).unwrap();
                        }
                        assert_eq!(
                            before,
                            Rc::as_ptr(&engines[1].render.borrow().stories["body"].blocks),
                            "paragraph {paragraph} offset {offset} {text:?} must patch"
                        );
                        let frames = frames();
                        assert_eq!(
                            frames[0], frames[1],
                            "paragraph {paragraph} offset {offset} {text:?}"
                        );
                    }
                }
            }
            for offset in (0..width).rev() {
                let paragraphs = engines[0].doc().paragraphs("body").unwrap();
                let start: u32 = paragraphs[..paragraph]
                    .iter()
                    .map(|paragraph| paragraph.text.encode_utf16().count() as u32 + 1)
                    .sum();
                let at = start + offset;
                let before = Rc::as_ptr(&engines[1].render.borrow().stories["body"].blocks);
                for engine in &engines {
                    engine
                        .edit_resident_text(StoryRange::new("body", at, at + 1), None, true)
                        .unwrap();
                    let epoch = engine.display.borrow().binary_frame_epoch;
                    engine.apply_and_layout("body", epoch).unwrap();
                }
                let after = Rc::as_ptr(&engines[1].render.borrow().stories["body"].blocks);
                assert_eq!(
                    before == after,
                    !(paragraph == 0 && matches!(offset, 2 | 4 | 6)),
                    "paragraph {paragraph} delete {offset}"
                );
                let frames = frames();
                assert_eq!(
                    frames[0], frames[1],
                    "paragraph {paragraph} delete {offset}"
                );
            }
        }
    }

    #[test]
    fn a_block_and_its_measure_fingerprint_as_their_measured_block() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let engine = paragraphs_engine(150, 40);
        engine
            .layout_document_with_regions_retained_json(&small_page_request(font_id))
            .unwrap();
        let pagination = engine.pagination.borrow();
        let measured = &pagination.input.as_ref().unwrap().measured;
        assert!(!measured.is_empty());
        for block in measured {
            assert_eq!(
                measured_parts_fingerprint(&block.block, &block.measure).unwrap(),
                measured_fingerprint(block).unwrap()
            );
        }
    }

    #[test]
    fn a_prefix_layout_lays_out_its_first_pages_like_the_full_pass() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let request = small_page_request(font_id);
        let engine = paragraphs_engine(142, 400);
        let full_json = paragraphs_engine(142, 400)
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let full: serde_json::Value = serde_json::from_str(&full_json).unwrap();

        let prefix: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap(),
        )
        .unwrap();
        let prefix_pages = prefix["layout"]["pages"].as_array().unwrap();
        let full_pages = full["layout"]["pages"].as_array().unwrap();
        assert_eq!(prefix["provisional"], true);
        assert!(full.get("provisional").is_none());
        assert!(prefix_pages.len() >= 5 && prefix_pages.len() < full_pages.len());
        assert_eq!(prefix_pages[..3], full_pages[..3]);
        {
            let pagination = engine.pagination.borrow();
            assert_eq!(
                pagination.measured_font_dependencies.len(),
                pagination.input.as_ref().unwrap().measured.len()
            );
            let chains = BTreeMap::from([("calibri|0|0".to_owned(), vec![font_id])]);
            assert!(
                pagination
                    .measured_font_dependencies
                    .iter()
                    .all(|dependencies| dependencies.matches(FontChains::BTree(&chains)))
            );
        }

        assert_eq!(
            engine
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
            full_json,
            "a full pass after a prefix pass matches a fresh one"
        );
        let short = paragraphs_engine(144, 3)
            .layout_document_with_regions_prefix_retained_json(&request, 3)
            .unwrap();
        assert!(
            !short.contains("provisional"),
            "a short body is laid out whole"
        );
    }

    #[test]
    fn a_prefix_with_floating_tables_lays_out_its_first_pages_like_the_full_pass() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let mut body = String::new();
        for index in 0..200 {
            if index % 7 == 3 {
                body.push_str(&format!(
                    r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="120" w:rightFromText="120" w:vertAnchor="text" w:horzAnchor="text" w:tblpY="60"/><w:tblW w:w="1800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid>{rows}</w:tbl>"#,
                    rows = format!("<w:tr><w:tc><w:p><w:r><w:t>Float {index}</w:t></w:r></w:p></w:tc></w:tr>").repeat(6)
                ));
            }
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index} wraps around the floating tables beside it on this page.</w:t></w:r></w:p>"
            ));
        }
        let bytes = docx_bytes("", &body);
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [{ "sectionId": "main", "properties": {
                "pageWidth": 5760, "pageHeight": 4320,
                "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
            } }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let seeded = || {
            let engine = EngineSession::new(148);
            crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
            engine
        };
        let full: serde_json::Value = serde_json::from_str(
            &seeded()
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
        )
        .unwrap();
        let engine = seeded();
        let prefix: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap(),
        )
        .unwrap();
        {
            let pagination = engine.pagination.borrow();
            assert_eq!(
                pagination.measured_font_dependencies.len(),
                pagination.input.as_ref().unwrap().measured.len()
            );
            let chains = BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]);
            assert!(
                pagination
                    .measured_font_dependencies
                    .iter()
                    .all(|dependencies| dependencies.matches(FontChains::BTree(&chains)))
            );
        }
        let first_pages = full["layout"]["pages"].as_array().unwrap()[..3].to_vec();
        assert!(
            serde_json::to_string(&first_pages)
                .unwrap()
                .contains("\"isFloating\":true")
        );
        assert_eq!(prefix["provisional"], true);
        assert_eq!(
            prefix["layout"]["pages"].as_array().unwrap()[..3],
            first_pages
        );
    }

    const INSIDE_SHAPE: &str = r#"<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="66675" distR="123825" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Inside shape"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;

    fn floating_table(anchor: &str, index: usize) -> String {
        format!(
            r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="120" w:rightFromText="120" {anchor} w:horzAnchor="text" w:tblpY="60"/><w:tblW w:w="1800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid>{rows}</w:tbl>"#,
            rows =
                format!("<w:tr><w:tc><w:p><w:r><w:t>Float {index}</w:t></w:r></w:p></w:tc></w:tr>")
                    .repeat(4)
        )
    }

    /// 200 paragraphs, with `float(index)` before those it returns one for.
    fn floated_body(float: impl Fn(usize) -> Option<String>) -> Vec<u8> {
        let mut body = String::new();
        for index in 0..200 {
            if let Some(float) = float(index) {
                body.push_str(&float);
            }
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index} wraps around the floats beside it on this page.</w:t></w:r></w:p>"
            ));
        }
        docx_bytes("", &body)
    }

    fn float_page_request(notes: serde_json::Value) -> String {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [{ "sectionId": "main", "properties": {
                "pageWidth": 5760, "pageHeight": 4320,
                "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
            } }] },
            "notes": notes,
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string()
    }

    /// The full pass and a three-page prefix pass of `bytes`, each on a fresh session
    /// that `prepare` sets up after seeding.
    fn full_and_prefix(
        bytes: &[u8],
        request: &str,
        prepare: &dyn Fn(&EngineSession),
    ) -> (serde_json::Value, serde_json::Value) {
        let seeded = || {
            let engine = EngineSession::new(149);
            crate::seed::seed_from_docx(engine.doc(), bytes).unwrap();
            prepare(&engine);
            engine
        };
        let full = seeded()
            .layout_document_with_regions_retained_json(request)
            .unwrap();
        let prefix = seeded()
            .layout_document_with_regions_prefix_retained_json(request, 3)
            .unwrap();
        (
            serde_json::from_str(&full).unwrap(),
            serde_json::from_str(&prefix).unwrap(),
        )
    }

    fn first_pages(output: &serde_json::Value) -> Vec<serde_json::Value> {
        output["layout"]["pages"].as_array().unwrap()[..3].to_vec()
    }

    #[test]
    fn floats_placed_from_outside_the_text_keep_the_whole_body() {
        let request = float_page_request(serde_json::json!({}));
        let off_page_shape = INSIDE_SHAPE
            .replace(
                r#"<wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH>"#,
                r#"<wp:positionH relativeFrom="outsideMargin"><wp:posOffset>-3175000</wp:posOffset></wp:positionH>"#,
            )
            .replace(
                r#"<wp:positionV relativeFrom="paragraph">"#,
                r#"<wp:positionV relativeFrom="margin">"#,
            )
            .replace(
                r#"<wp:wrapSquare wrapText="bothSides"/>"#,
                "<wp:wrapTopAndBottom/>",
            );
        assert_ne!(off_page_shape, INSIDE_SHAPE);
        for (name, float, every) in [
            (
                "margin table",
                floating_table(r#"w:vertAnchor="margin""#, 3),
                7,
            ),
            ("page table", floating_table(r#"w:vertAnchor="page""#, 3), 7),
            ("table without an anchor", floating_table("", 3), 7),
            ("inside shape", format!("<w:p>{INSIDE_SHAPE}</w:p>"), 7),
            // One late shape that starts off the page, so it has no zone until
            // page-side wrapping brings it in.
            (
                "late off-page shape",
                format!("<w:p>{off_page_shape}</w:p>"),
                150,
            ),
        ] {
            let bytes = floated_body(|index| (index % every == every - 1).then(|| float.clone()));
            let (full, prefix) = full_and_prefix(&bytes, &request, &|_| {});
            assert_ne!(prefix["provisional"], true, "{name}");
            assert_eq!(prefix["layout"], full["layout"], "{name}");
        }
    }

    #[test]
    fn a_float_next_to_the_cut_leaves_the_first_pages_as_the_full_pass_lays_them_out() {
        let request = float_page_request(serde_json::json!({}));
        let bytes = floated_body(|index| {
            (index % 2 == 1).then(|| floating_table(r#"w:vertAnchor="text""#, index))
        });
        let (full, prefix) = full_and_prefix(&bytes, &request, &|_| {});
        assert_eq!(prefix["provisional"], true);
        assert_eq!(first_pages(&prefix), first_pages(&full));
    }

    #[test]
    fn a_float_beside_a_footnote_leaves_the_first_pages_as_the_full_pass_lays_them_out() {
        let request = float_page_request(serde_json::json!({"contents": [{"id": 5, "height": 0}]}));
        let bytes = floated_body(|index| {
            (index % 3 == 0).then(|| floating_table(r#"w:vertAnchor="text""#, index))
        });
        let with_note = |engine: &EngineSession| {
            engine
                .doc()
                .insert_embed(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", 20),
                    "noteRef",
                    vec![("footnoteRefId".to_owned(), Any::Number(5.0))],
                )
                .unwrap();
            engine
                .doc()
                .create_story("fn:5", "A footnote beside the floats", "Normal", "left")
                .unwrap();
        };
        let (full, prefix) = full_and_prefix(&bytes, &request, &with_note);
        assert_eq!(prefix["provisional"], true);
        assert!(
            serde_json::to_string(&first_pages(&full))
                .unwrap()
                .contains("footnoteIds"),
            "the note sits on the first pages"
        );
        assert_eq!(first_pages(&prefix), first_pages(&full));
    }

    #[test]
    fn a_provisional_layout_renders_numpages_empty_until_the_full_pass() {
        assert_partial_numpages(false);
    }

    #[test]
    fn a_provisional_layout_renders_cached_numpages_until_the_full_pass() {
        assert_partial_numpages(true);
    }

    fn assert_partial_numpages(cached_page_totals: bool) {
        let mut request = float_page_request(serde_json::json!({}));
        if cached_page_totals {
            let mut value: serde_json::Value = serde_json::from_str(&request).unwrap();
            value["cachedPageTotals"] = serde_json::json!(true);
            request = value.to_string();
        }
        let mut body = String::from(
            r#"<w:p><w:r><w:t>Page count </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>"#,
        );
        for index in 0..200 {
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index} fills the pages the field counts.</w:t></w:r></w:p>"
            ));
        }
        let bytes = docx_bytes("", &body);
        let engine = EngineSession::new(150);
        crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
        let numpages_text = |engine: &EngineSession| {
            engine.build_display_list_frame("{}", 0).unwrap();
            engine
                .with_display_list(|list| {
                    let page = serde_json::to_value(&list.pages[0]).unwrap();
                    page["primitives"]
                        .as_array()
                        .unwrap()
                        .iter()
                        .filter(|primitive| primitive["field"]["category"] == "NUMPAGES")
                        .map(|primitive| primitive["text"].as_str().unwrap_or("").to_owned())
                        .collect::<String>()
                })
                .unwrap()
        };
        let prefix: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(prefix["provisional"], true);
        assert_eq!(prefix["layout"]["partial"], true);
        assert_eq!(
            prefix["layout"]["cachedPageTotals"],
            if cached_page_totals {
                serde_json::json!(true)
            } else {
                serde_json::Value::Null
            }
        );
        assert_eq!(
            numpages_text(&engine),
            if cached_page_totals { "9" } else { "" }
        );

        let full: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
        )
        .unwrap();
        let pages = full["layout"]["pages"].as_array().unwrap().len();
        assert!(pages > prefix["layout"]["pages"].as_array().unwrap().len());
        assert_eq!(numpages_text(&engine), pages.to_string());

        engine.set_partial_document(true);
        let cut: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(cut["layout"]["partial"], true);
        assert_eq!(
            numpages_text(&engine),
            if cached_page_totals { "9" } else { "" }
        );

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 2),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        assert!(
            engine
                .apply_and_layout_regions_resident("body", &mut |_| {})
                .unwrap()
        );
        assert_eq!(
            engine
                .pagination
                .borrow()
                .layout
                .as_ref()
                .unwrap()
                .cached_page_totals,
            cached_page_totals
        );
        assert_eq!(
            numpages_text(&engine),
            if cached_page_totals { "9" } else { "" }
        );

        if cached_page_totals {
            let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
            request.as_object_mut().unwrap().remove("cachedPageTotals");
            let request = request.to_string();
            let default: serde_json::Value = serde_json::from_str(
                &engine
                    .layout_document_with_regions_retained_json(&request)
                    .unwrap(),
            )
            .unwrap();
            assert!(default["layout"].get("cachedPageTotals").is_none());
            assert_eq!(numpages_text(&engine), "");
        }
    }

    #[test]
    fn a_prefix_ends_after_a_keep_with_next_run_and_its_follower() {
        let paragraph = |keep_next: bool| {
            let mut block: LayoutBlock = serde_json::from_value(serde_json::json!({
                "kind": "paragraph", "id": "p", "runs": []
            }))
            .unwrap();
            if let LayoutBlock::Paragraph(paragraph) = &mut block {
                paragraph.attrs = Some(docx_layout::types::ParagraphAttrs {
                    keep_next: Some(keep_next),
                    ..Default::default()
                });
            }
            block
        };
        let blocks = [
            paragraph(false),
            paragraph(true),
            paragraph(true),
            paragraph(false),
            paragraph(false),
        ];
        assert_eq!(prefix_boundary(&blocks, 1), 1);
        assert_eq!(prefix_boundary(&blocks, 2), 4);
        assert_eq!(prefix_boundary(&blocks, 3), 4);
    }

    #[test]
    fn a_prefix_cut_inside_an_earlier_section_lays_it_out_with_its_own_geometry() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let paragraph = |index: usize| {
            format!("<w:p><w:r><w:t>Paragraph {index} in the first section.</w:t></w:r></w:p>")
        };
        let mut body: String = (0..160).map(paragraph).collect();
        body.push_str(
            r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="4320" w:h="2880"/></w:sectPr></w:pPr></w:p>"#,
        );
        body.extend((160..200).map(paragraph));
        let bytes = docx_bytes("", &body);
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [
                { "sectionId": "first", "properties": {
                    "pageWidth": 4320, "pageHeight": 2880,
                    "marginTop": 300, "marginRight": 300, "marginBottom": 300, "marginLeft": 300
                } },
                { "sectionId": "last", "properties": {
                    "pageWidth": 6480, "pageHeight": 1440,
                    "marginTop": 200, "marginRight": 200, "marginBottom": 200, "marginLeft": 200
                } }
            ] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let seeded = |client_id| {
            let engine = EngineSession::new(client_id);
            crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
            engine
        };
        let full: serde_json::Value = serde_json::from_str(
            &seeded(147)
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
        )
        .unwrap();
        let prefix: serde_json::Value = serde_json::from_str(
            &seeded(147)
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(prefix["provisional"], true);
        assert_eq!(
            prefix["layout"]["pages"].as_array().unwrap()[..3],
            full["layout"]["pages"].as_array().unwrap()[..3]
        );
    }

    #[test]
    fn a_float_past_a_narrow_first_section_lays_out_the_first_pages_like_the_full_pass() {
        // 200 px in: past the first section's 96 px column, within the second's.
        let shape = INSIDE_SHAPE
            .replace(
                r#"<wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH>"#,
                r#"<wp:positionH relativeFrom="margin"><wp:posOffset>1905000</wp:posOffset></wp:positionH>"#,
            )
            .replace(
                r#"<wp:wrapSquare wrapText="bothSides"/>"#,
                "<wp:wrapTopAndBottom/>",
            );
        let mut body = String::new();
        for index in 0..400 {
            if index == 10 {
                body.push_str(
                    r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="2160" w:h="8000"/></w:sectPr></w:pPr></w:p>"#,
                );
            }
            if index == 40 {
                body.push_str(&format!("<w:p>{shape}</w:p>"));
            }
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"
            ));
        }
        let bytes = docx_bytes("", &body);
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let section = |id: &str, width: u32| {
            serde_json::json!({ "sectionId": id, "properties": {
                "pageWidth": width, "pageHeight": 8000,
                "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
            } })
        };
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [section("first", 2160), section("last", 5760)] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let (full, prefix) = full_and_prefix(&bytes, &request, &|_| {});
        assert_eq!(prefix["provisional"], true);
        assert_eq!(first_pages(&prefix), first_pages(&full));
    }

    #[test]
    fn a_prefix_does_not_end_inside_a_section_with_columns() {
        let mut request: serde_json::Value =
            serde_json::from_str(&float_page_request(serde_json::json!({}))).unwrap();
        let properties = &mut request["regions"]["sections"][0]["properties"];
        properties["pageWidth"] = 12240.into();
        properties["pageHeight"] = 15840.into();
        properties["columnCount"] = 2.into();
        properties["columnSpace"] = 300.into();
        let request = request.to_string();
        // Balanced over 32 paragraphs, the first column would end sooner than
        // over all 100, which are too tall to balance.
        let mut body = format!(
            "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            "A long first paragraph of text and more text. ".repeat(100)
        );
        for index in 1..100 {
            body.push_str(&format!(
                "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"
            ));
        }
        body.push_str(r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360" w:header="720" w:footer="720"/><w:cols w:num="2" w:space="300"/></w:sectPr>"#);
        let (full, prefix) = full_and_prefix(&docx_bytes("", &body), &request, &|_| {});
        assert_eq!(prefix["layout"], full["layout"]);
    }

    #[test]
    fn an_edit_over_a_prefix_layout_lays_out_the_whole_body() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(FONT).unwrap();
        let request = small_page_request(font_id);
        let extras = serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string();
        let edit = |engine: &EngineSession| {
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", 4),
                    "x",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
        };
        let engine = paragraphs_engine(145, 400);
        engine
            .layout_document_with_regions_prefix_retained_json(&request, 3)
            .unwrap();
        engine.build_display_list_frame(&extras, 0).unwrap();
        edit(&engine);
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();

        let reference = paragraphs_engine(145, 400);
        edit(&reference);
        reference
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let pages = |engine: &EngineSession| {
            serde_json::to_value(&engine.pagination.borrow().layout.as_ref().unwrap().pages)
                .unwrap()
        };
        assert_eq!(pages(&engine), pages(&reference));
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
    fn region_font_preflight_covers_generated_note_labels() {
        for (prefix, note_kind) in [("fn", "footnote"), ("en", "endnote")] {
            let engine = EngineSession::new(1371);
            let body = [serde_json::json!({
                "type": "paragraph",
                "formatting": {"runProperties": {"fontFamily": {"ascii": "Arial"}}},
                "content": [{
                    "type": "run",
                    "formatting": {"fontFamily": {
                        "ascii": "Arial", "hAnsi": "Arial", "cs": "Calibri"
                    }},
                    "content": [{"type": "text", "text": "Body"}]
                }]
            })];
            let note = [serde_json::json!({
                "type": "paragraph",
                "formatting": {"runProperties": {"fontFamily": {"ascii": "Arial"}}},
                "content": []
            })];
            crate::seed::seed_blocks(
                engine.doc(),
                None,
                &[("body".to_owned(), &body), (format!("{prefix}:5"), &note)],
            )
            .unwrap();
            for preview in [serde_json::json!({}), serde_json::json!({"9": "accepted"})] {
                let request = serde_json::json!({
                    "bodyStory": "body",
                    "notes": {"contents": [{"id": 5, "noteKind": note_kind, "height": 0}]},
                    "measurement": {"defaults": {"fontFamily": "Calibri"}},
                    "renderEnv": {"revisionPreview": preview}
                });
                let requirements: Vec<serde_json::Value> = serde_json::from_str(
                    &engine
                        .layout_font_requirements_json(&request.to_string())
                        .unwrap(),
                )
                .unwrap();

                assert!(
                    requirements
                        .iter()
                        .any(|requirement| requirement["key"] == "calibri|0|0"),
                    "{note_kind}: {requirements:?}"
                );
            }
        }
    }

    #[test]
    fn preview_font_preflight_reuses_markup_until_the_document_changes() {
        let engine = EngineSession::new(1361);
        crate::seed::seed_from_docx(
            engine.doc(),
            &docx_bytes(
                "",
                r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:t>Beta</w:t></w:r></w:p>"#,
            ),
        )
        .unwrap();
        let deletion = engine
            .doc()
            .delete_range(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::StoryRange::new("body", 0, 5),
            )
            .unwrap();
        let id = &deletion.revision_ids[0];
        let mut request = serde_json::json!({"bodyStory": "body", "renderEnv": {}});
        let markup = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        request["renderEnv"]["revisionPreview"] = serde_json::json!({id: "accepted"});
        let first = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        assert_eq!(first, markup);
        let requirements: Vec<serde_json::Value> = serde_json::from_str(&first).unwrap();
        assert!(
            requirements
                .iter()
                .any(|requirement| requirement["key"] == "courier new|0|0")
        );

        let before = engine.stats();
        request["renderEnv"]["revisionPreview"] = serde_json::json!({id: "rejected"});
        assert_eq!(
            engine
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
            first
        );
        assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
        assert_eq!(engine.stats().lower_cache_hits, before.lower_cache_hits);

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 0),
                "New",
                crate::FormatPolicy::Explicit(BTreeMap::from([(
                    "fontFamily".to_owned(),
                    Any::from("Times New Roman"),
                )])),
            )
            .unwrap();
        assert_ne!(engine.doc_epoch(), before.doc_epoch);
        let updated = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        assert_ne!(updated, first);
        let requirements: Vec<serde_json::Value> = serde_json::from_str(&updated).unwrap();
        assert!(
            requirements
                .iter()
                .any(|requirement| requirement["key"] == "times new roman|0|0")
        );
        assert_eq!(
            engine.stats().lower_cache_misses,
            before.lower_cache_misses + 1
        );
        request["renderEnv"] = serde_json::json!({});
        assert_eq!(
            updated,
            engine
                .layout_font_requirements_json(&request.to_string())
                .unwrap()
        );
    }

    #[test]
    fn preview_font_preflight_keys_the_remaining_request() {
        let engine = EngineSession::new(1362);
        engine
            .doc()
            .create_story("body", "Text", "Normal", "left")
            .unwrap();
        let mut request = serde_json::json!({
            "bodyStory": "body",
            "renderEnv": {"revisionPreview": {"id": "accepted"}},
            "measurement": {"defaults": {"fontFamily": "Calibri"}}
        });
        let first = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        request["measurement"]["defaults"]["fontFamily"] = "Courier New".into();
        let second = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        assert_ne!(second, first);
        assert_eq!(
            engine
                .preview_font_requirements
                .borrow()
                .as_ref()
                .unwrap()
                .json
                .as_deref(),
            Some(second.as_str())
        );
    }

    #[test]
    fn preview_font_preflight_covers_a_hidden_list_marker_before_a_page_break() {
        let engine = EngineSession::new(1363);
        let para_id = engine
            .doc()
            .create_story("body", "", "Normal", "left")
            .unwrap();
        for (key, value) in [
            ("numPr", Any::from_json(r#"{"numId":1,"ilvl":0}"#).unwrap()),
            ("listMarkerFontFamily", Any::from("Courier New")),
            ("listMarkerBold", Any::Bool(true)),
        ] {
            engine
                .doc()
                .set_paragraph_attr(&para_id, key, value)
                .unwrap();
        }
        let page_break = engine
            .doc()
            .insert_embed(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::Position::new("body", 1),
                "pageBreak",
                Vec::new(),
            )
            .unwrap();
        let markup = engine
            .layout_font_requirements_json(r#"{"bodyStory":"body","renderEnv":{}}"#)
            .unwrap();
        let requirements: Vec<serde_json::Value> = serde_json::from_str(&markup).unwrap();
        assert!(
            !requirements
                .iter()
                .any(|requirement| requirement["key"] == "courier new|1|0")
        );
        for render_env in [
            serde_json::json!({"revisionPreview": {}}),
            serde_json::json!({"revisionPreview": {"id": "proposed"}}),
        ] {
            assert_eq!(
                engine
                    .layout_font_requirements_json(
                        &serde_json::json!({"bodyStory": "body", "renderEnv": render_env})
                            .to_string(),
                    )
                    .unwrap(),
                markup
            );
        }

        let id = &page_break.revision_ids[0];
        let request = serde_json::json!({
            "bodyStory": "body",
            "renderEnv": {"revisionPreview": {id: "rejected"}}
        });
        let env: RenderEnv = serde_json::from_value(request["renderEnv"].clone()).unwrap();
        let blocks = crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap();
        assert!(
            !blocks
                .iter()
                .any(|block| matches!(block, LayoutBlock::PageBreak(_)))
        );
        let exact = docx_layout::measure_blocks::collect_font_requirements(&blocks, "Calibri");
        assert!(
            exact
                .iter()
                .any(|requirement| requirement.key == "courier new|1|0")
        );
        let superset = engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap();
        let requirements: Vec<serde_json::Value> = serde_json::from_str(&superset).unwrap();
        for requirement in exact {
            assert!(requirements.contains(&serde_json::to_value(requirement).unwrap()));
        }
    }

    fn font_preflight_run(text: &str, family: &str) -> serde_json::Value {
        serde_json::json!({
            "type": "run",
            "formatting": {"fontFamily": {"ascii": family, "hAnsi": family}},
            "content": [{"type": "text", "text": text}]
        })
    }

    fn assert_preview_font_preflight_covers(
        engine: &EngineSession,
        revision_id: &str,
        decision: crate::bridge::RevisionPreview,
    ) -> Vec<docx_layout::measure_blocks::FontRequirement> {
        let env = RenderEnv::default().with_revision_preview(revision_id, decision);
        let blocks = crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap();
        let exact = docx_layout::measure_blocks::collect_font_requirements(&blocks, "Calibri");
        let request = serde_json::json!({"bodyStory": "body", "renderEnv": env});
        let superset: Vec<serde_json::Value> = serde_json::from_str(
            &engine
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        for requirement in &exact {
            let candidate = superset
                .iter()
                .find(|candidate| candidate["key"] == requirement.key)
                .unwrap_or_else(|| panic!("missing font requirement: {requirement:?}"));
            assert_eq!(candidate["family"], requirement.family);
            assert_eq!(candidate["bold"], requirement.bold);
            assert_eq!(candidate["italic"], requirement.italic);
            for script in &requirement.scripts {
                assert!(
                    candidate["scripts"]
                        .as_array()
                        .is_some_and(|scripts| scripts.iter().any(|value| value == script)),
                    "missing script {script} for {}: {candidate}",
                    requirement.key
                );
            }
        }
        exact
    }

    #[test]
    fn preview_font_preflight_covers_a_surviving_list_marker_family() {
        let engine = EngineSession::new(1364);
        let blocks = [serde_json::json!({
            "type": "paragraph",
            "formatting": {"numPr": {"numId": 1, "ilvl": 0}},
            "listRendering": {"marker": "%1.", "markerBold": true},
            "content": [
                font_preflight_run("First", "Courier New"),
                font_preflight_run("Second", "Times New Roman")
            ]
        })];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let deletion = engine
            .doc()
            .delete_range(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::StoryRange::new("body", 0, 5),
            )
            .unwrap();
        let exact = assert_preview_font_preflight_covers(
            &engine,
            &deletion.revision_ids[0],
            crate::bridge::RevisionPreview::Accepted,
        );
        assert!(
            exact
                .iter()
                .any(|requirement| requirement.key == "times new roman|1|0")
        );
    }

    #[test]
    fn preview_font_preflight_covers_an_unrenderable_list_marker_after_renumbering() {
        let engine = EngineSession::new(1365);
        let item = |family| {
            serde_json::json!({
                "type": "paragraph",
                "formatting": {"numPr": {"numId": 1, "ilvl": 0}},
                "listRendering": {
                    "marker": "%1.", "levelNumFmts": ["upperRoman"],
                    "startOverride": 3999, "markerBold": true, "markerItalic": true
                },
                "content": [font_preflight_run("Item", family)]
            })
        };
        let blocks = [
            serde_json::json!({
                "type": "table",
                "rows": [{"type": "tableRow", "cells": [{
                    "type": "tableCell", "content": [item("Calibri")]
                }]}]
            }),
            item("Preview Roman"),
        ];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let markup =
            crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default())
                .unwrap();
        let LayoutBlock::Paragraph(paragraph) = &markup[1] else {
            panic!("expected the second list item");
        };
        assert!(paragraph.attrs.as_ref().unwrap().list_marker.is_none());
        let deletion = engine
            .doc()
            .delete_range(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::StoryRange::new("body", 0, 1),
            )
            .unwrap();
        let exact = assert_preview_font_preflight_covers(
            &engine,
            &deletion.revision_ids[0],
            crate::bridge::RevisionPreview::Accepted,
        );
        assert!(
            exact
                .iter()
                .any(|requirement| requirement.key == "preview roman|1|1")
        );
    }

    #[test]
    fn preview_font_preflight_covers_han_without_an_inserted_kana_run() {
        let engine = EngineSession::new(1366);
        let blocks = [serde_json::json!({
            "type": "paragraph", "content": [font_preflight_run("漢", "Preview Han")]
        })];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        let insertion = engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::Position::new("body", 1),
                "かな",
                crate::FormatPolicy::Explicit(BTreeMap::from([(
                    "fontFamily".to_owned(),
                    Any::from("Preview Kana"),
                )])),
            )
            .unwrap();
        let exact = assert_preview_font_preflight_covers(
            &engine,
            &insertion.revision_ids[0],
            crate::bridge::RevisionPreview::Rejected,
        );
        assert!(exact.iter().any(|requirement| {
            requirement.key == "preview han|0|0"
                && requirement.scripts.contains(&"cjk-sc".to_owned())
        }));
    }

    #[test]
    fn preview_font_preflight_takes_the_exact_path_for_script_fallbacks() {
        let cases: [(u64, &str, &str, &[&str]); 2] = [
            (1371, "骨", "Calibri", &["cjk-sc", "cjk-jp"]),
            (1372, "a😀b", "Arial", &[]),
        ];
        for (client_id, text, family, scripts) in cases {
            let engine = EngineSession::new(client_id);
            let blocks = [serde_json::json!({
                "type": "paragraph", "content": [font_preflight_run(text, family)]
            })];
            crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
            let insertion = engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                    crate::Position::new("body", 1),
                    "かな",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            for (index, decision) in [
                crate::bridge::RevisionPreview::Rejected,
                crate::bridge::RevisionPreview::Accepted,
                crate::bridge::RevisionPreview::Rejected,
            ]
            .into_iter()
            .enumerate()
            {
                let env = RenderEnv::default()
                    .with_revision_preview(&insertion.revision_ids[0], decision);
                let blocks =
                    crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap();
                let exact = docx_layout::measure_blocks::collect_font_requirements(&blocks, family);
                let request = serde_json::json!({
                    "bodyStory": "body",
                    "renderEnv": env,
                    "measurement": {"defaults": {"fontFamily": family}}
                });
                let requirements: serde_json::Value = serde_json::from_str(
                    &engine
                        .layout_font_requirements_json(&request.to_string())
                        .unwrap(),
                )
                .unwrap();
                assert_eq!(requirements, serde_json::to_value(exact).unwrap());
                if let Some(script) = scripts.get(index % 2) {
                    assert_eq!(requirements[0]["scripts"], serde_json::json!([script]));
                }
                assert!(
                    engine
                        .preview_font_requirements
                        .borrow()
                        .as_ref()
                        .is_some_and(|cached| cached.json.is_none())
                );
            }
        }
    }

    #[test]
    fn preview_font_preflight_covers_suppressed_numeric_field_results() {
        assert_preview_font_preflight_covers_numeric_field_results(1367, false);
    }

    #[test]
    fn preview_font_preflight_covers_suppressed_results_of_a_hidden_numeric_field() {
        assert_preview_font_preflight_covers_numeric_field_results(1368, true);
    }

    fn assert_preview_font_preflight_covers_numeric_field_results(seed: u64, hidden: bool) {
        let engine = EngineSession::new(seed);
        let cached = serde_json::json!({
            "type": "paragraph", "content": [font_preflight_run("漢", "Preview Field")]
        });
        let table = serde_json::json!({
            "type": "table", "rows": [{"type": "tableRow", "cells": [{
                "type": "tableCell", "content": [{
                    "type": "paragraph",
                    "content": [font_preflight_run("Table", "Preview Field Table")]
                }]
            }]}]
        });
        let sdt = serde_json::json!({
            "type": "blockSdt", "properties": {}, "content": [{
                "type": "paragraph",
                "content": [font_preflight_run("SDT", "Preview Field SDT")]
            }]
        });
        let end = serde_json::json!({"type": "paragraph", "content": []});
        let field = serde_json::json!({
            "type": "complexField", "fieldType": "UNKNOWN", "instruction": "0",
            "fieldCode": [], "fieldResult": [font_preflight_run("First", "Calibri")],
            "structuredResult": {"blocks": [cached, table, sdt, end]}
        });
        let blocks = [
            serde_json::json!({"type": "paragraph", "content": [field]}),
            cached,
            table,
            sdt,
            end,
        ];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![crate::RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: Attrs::from_iter(
                        [(
                            "ins".into(),
                            Any::from_json(
                                r#"{"id":"9","author":"Ann","date":"2026-09-29T12:00:00Z"}"#,
                            )
                            .unwrap(),
                        )]
                        .into_iter()
                        .chain(hidden.then(|| ("hidden".into(), Any::Bool(true)))),
                    ),
                }],
                &crate::EditCtx::local("", ""),
            )
            .unwrap();
        let markup: Vec<serde_json::Value> = serde_json::from_str(
            &engine
                .layout_font_requirements_json(r#"{"bodyStory":"body","renderEnv":{}}"#)
                .unwrap(),
        )
        .unwrap();
        assert!(markup.iter().all(|requirement| {
            !requirement["family"]
                .as_str()
                .unwrap()
                .starts_with("Preview")
        }));
        let exact = assert_preview_font_preflight_covers(
            &engine,
            "9",
            crate::bridge::RevisionPreview::Rejected,
        );
        for key in [
            "preview field|0|0",
            "preview field table|0|0",
            "preview field sdt|0|0",
        ] {
            assert!(exact.iter().any(|requirement| requirement.key == key));
        }
    }

    fn insert_tracked_text_shape(engine: &EngineSession, index: u32, revision_id: &str) {
        let shape = serde_json::json!({
            "shapeType": "rect",
            "size": {"width": 914400, "height": 457200},
            "textBody": {"content": [{
                "paraId": "p1",
                "content": [{"type": "run", "content": [{"type": "text", "text": "hi"}]}]
            }]}
        });
        let revision = serde_json::json!({
            "id": revision_id, "author": "Ann", "date": "2026-09-29T12:00:00Z"
        });
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![crate::RawOp::InsertEmbed {
                    index,
                    kind: "shape".to_owned(),
                    payload: vec![("shapeJson".to_owned(), Any::from(shape.to_string()))],
                    attrs: Attrs::from([(
                        "ins".into(),
                        Any::from_json(&revision.to_string()).unwrap(),
                    )]),
                }],
                &crate::EditCtx::local("", ""),
            )
            .unwrap();
    }

    #[test]
    fn preview_font_preflight_covers_segments_a_hidden_drawing_joins() {
        let engine = EngineSession::new(1369);
        let blocks = [serde_json::json!({
            "type": "paragraph",
            "content": [
                font_preflight_run("A", "Courier New"),
                font_preflight_run("漢", "Preview Han")
            ]
        })];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        insert_tracked_text_shape(&engine, 1, "7");
        let markup =
            crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default())
                .unwrap();
        assert!(matches!(markup[1], LayoutBlock::Shape(_)));
        let exact = assert_preview_font_preflight_covers(
            &engine,
            "7",
            crate::bridge::RevisionPreview::Rejected,
        );
        assert!(exact.iter().any(|requirement| {
            requirement.key == "courier new|0|0"
                && requirement.scripts.contains(&"cjk-sc".to_owned())
        }));
    }

    #[test]
    fn preview_font_preflight_covers_a_paragraph_a_hidden_drawing_empties() {
        let engine = EngineSession::new(1370);
        let blocks = [serde_json::json!({
            "type": "paragraph",
            "formatting": {"numPr": {"numId": 1, "ilvl": 0}},
            "listRendering": {
                "marker": "%1.", "markerBold": true, "markerFontFamily": "Preview Marker"
            },
            "content": []
        })];
        crate::seed::seed_blocks(engine.doc(), None, &[("body".to_owned(), &blocks)]).unwrap();
        insert_tracked_text_shape(&engine, 0, "8");
        let markup =
            crate::bridge::yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default())
                .unwrap();
        assert!(
            !markup
                .iter()
                .any(|block| matches!(block, LayoutBlock::Paragraph(_)))
        );
        let exact = assert_preview_font_preflight_covers(
            &engine,
            "8",
            crate::bridge::RevisionPreview::Rejected,
        );
        assert!(
            exact
                .iter()
                .any(|requirement| requirement.key == "preview marker|1|0")
        );
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
        let area = &output["layout"]["pages"][0]["noteAreas"][0];
        let note = &area["notes"][0];
        let separator_height = area["separator"]["height"]
            .as_f64()
            .unwrap_or(docx_layout::footnotes::FOOTNOTE_SEPARATOR_HEIGHT);

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
                > separator_height
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
                .resident_layout_input_from_blocks(blocks, false, false, &mut |index, key, _, _| {
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
    fn recovery_frames_are_newer_than_the_frame_the_caller_holds() {
        let engine = EngineSession::new(18);
        engine
            .layout_document_json(
                r#"{"measured": [], "options": {"pageSize": {"w": 816, "h": 1056},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}}"#,
            )
            .unwrap();
        engine.build_display_list_frame("{}", 7).unwrap();
        assert_eq!(engine.display.borrow().binary_frame_epoch, 8);
        engine.build_display_list_frame("{}", 8).unwrap();
        assert_eq!(engine.display.borrow().binary_frame_epoch, 9);
        engine.build_display_list_frame("{}", 3).unwrap();
        assert_eq!(engine.display.borrow().binary_frame_epoch, 10);
        engine.reset_frame_base();
        let full = engine.build_display_list_frame("{}", 10).unwrap();
        assert_eq!(u32::from_le_bytes(full[12..16].try_into().unwrap()), 1);
        let delta = engine.build_display_list_frame("{}", 11).unwrap();
        assert_eq!(u32::from_le_bytes(delta[12..16].try_into().unwrap()), 0);
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

    fn layout_with_note_separator(
        kind: &str,
        separator: Option<&str>,
    ) -> (EngineSession, serde_json::Value) {
        let mut body = format!(
            r#"<w:p><w:r><w:t>BODY 01</w:t></w:r><w:r><w:{kind}Reference w:id="1"/></w:r></w:p>"#
        );
        for index in 1..=54 {
            body.push_str(&format!(
                "<w:p><w:r><w:t>FILLER {index:02}</w:t></w:r></w:p>"
            ));
        }
        body.push_str(r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#);
        layout_note_separator_document(
            kind,
            separator,
            &body,
            serde_json::json!([{"sectionId": "main", "properties": {
                "pageWidth": 11906, "pageHeight": 16838,
                "marginTop": 1440, "marginRight": 1440, "marginBottom": 1440, "marginLeft": 1440
            }}]),
            &[1],
        )
    }

    fn layout_note_separator_document(
        kind: &str,
        separator: Option<&str>,
        body: &str,
        sections: serde_json::Value,
        note_ids: &[i64],
    ) -> (EngineSession, serde_json::Value) {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Separator"><w:name w:val="Separator"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>"#;
        let separator = separator
            .map(|paragraphs| {
                format!(r#"<w:{kind} w:id="-1" w:type="separator">{paragraphs}</w:{kind}>"#)
            })
            .unwrap_or_default();
        let note_content: String = note_ids
            .iter()
            .map(|id| format!(r#"<w:{kind} w:id="{id}"><w:p><w:r><w:{kind}Ref/></w:r><w:r><w:t>NOTE {id:02}</w:t></w:r></w:p></w:{kind}>"#))
            .collect();
        let notes = format!(
            r#"<w:{kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{separator}<w:{kind} w:id="0" w:type="continuationSeparator"><w:p><w:pPr><w:spacing w:line="1920" w:lineRule="exact"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>{note_content}</w:{kind}s>"#
        );
        let mut parts = ooxml_opc::unzip_parts(&docx_bytes(styles, body)).unwrap();
        for (path, bytes) in &mut parts {
            let addition = match path.as_str() {
                "[Content_Types].xml" => Some((
                    "</Types>",
                    format!(
                        r#"<Override PartName="/word/{kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}s+xml"/></Types>"#
                    ),
                )),
                "word/_rels/document.xml.rels" => Some((
                    "</Relationships>",
                    format!(
                        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}s" Target="{kind}s.xml"/></Relationships>"#
                    ),
                )),
                _ => None,
            };
            if let Some((closing, replacement)) = addition {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace(closing, &replacement)
                    .into_bytes();
            }
        }
        parts.push((format!("word/{kind}s.xml"), notes.into_bytes()));
        let engine = EngineSession::new(312);
        crate::seed::seed_from_docx(engine.doc(), &ooxml_opc::rezip_parts(&parts).unwrap())
            .unwrap();
        let contents: Vec<_> = note_ids
            .iter()
            .map(|id| serde_json::json!({"id": id, "noteKind": kind, "height": 0}))
            .collect();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": sections},
            "notes": {"contents": contents},
            "measurement": {
                "fontChains": {"liberation sans|0|0": [font_id]},
                "defaults": {"fontSize": 12, "fontFamily": "Liberation Sans"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        let mut output: serde_json::Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        (engine, output["layout"].take())
    }

    #[test]
    fn a_twelve_point_separator_keeps_body_lines_on_one_page() {
        let (_, layout) = layout_with_note_separator(
            "footnote",
            Some(
                r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#,
            ),
        );
        assert_eq!(layout["pages"].as_array().unwrap().len(), 1);
    }

    #[test]
    fn a_taller_separator_moves_body_lines_to_the_next_page() {
        let (engine, _) = layout_with_note_separator(
            "footnote",
            Some(
                r#"<w:p><w:pPr><w:spacing w:line="960" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#,
            ),
        );
        let pagination = engine.pagination.borrow();
        let input = pagination.input.as_ref().unwrap();
        let layout = pagination.layout.as_ref().unwrap();
        let last_page_text: Vec<_> = layout
            .pages
            .last()
            .unwrap()
            .fragments
            .iter()
            .filter_map(|fragment| {
                let Fragment::Paragraph(fragment) = fragment else {
                    return None;
                };
                input.measured.iter().find_map(|measured| {
                    let LayoutBlock::Paragraph(block) = &measured.block else {
                        return None;
                    };
                    (block.id == fragment.block_id).then(|| {
                        block
                            .runs
                            .iter()
                            .filter_map(|run| match run {
                                Run::Text(text) => Some(text.text.as_str()),
                                _ => None,
                            })
                            .collect::<String>()
                    })
                })
            })
            .collect();
        assert_eq!(
            (layout.pages.len(), last_page_text),
            (2, vec!["FILLER 53".to_owned(), "FILLER 54".to_owned()])
        );
    }

    #[test]
    fn replicated_separator_state_preserves_note_heights_and_body_breaks() {
        let page_breaks = |layout: &serde_json::Value| {
            layout["pages"]
                .as_array()
                .unwrap()
                .iter()
                .map(|page| {
                    (
                        page["fragments"].clone(),
                        page["footnoteReservedHeight"].clone(),
                        page["noteAreas"]
                            .as_array()
                            .into_iter()
                            .flatten()
                            .map(|area| area["separator"]["height"].clone())
                            .collect::<Vec<_>>(),
                    )
                })
                .collect::<Vec<_>>()
        };
        for kind in ["footnote", "endnote"] {
            for line in [960, 1440] {
                let separator = format!(
                    r#"<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#
                );
                let (opener, layout) = layout_with_note_separator(kind, Some(&separator));
                let request = opener
                    .regions
                    .borrow()
                    .as_ref()
                    .unwrap()
                    .request_json
                    .clone();
                let replica = EngineSession::new(313);
                replica
                    .doc()
                    .apply_update_v1(&opener.doc().encode_state_as_update_v1())
                    .unwrap();
                assert!(replica.doc().source_metadata().is_none());
                let fallback: serde_json::Value = serde_json::from_str(
                    &replica.layout_document_with_regions_json(&request).unwrap(),
                )
                .unwrap();
                for area in fallback["layout"]["pages"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .flat_map(|page| page["noteAreas"].as_array().into_iter().flatten())
                {
                    assert!(area.get("separator").is_none());
                    let notes_height: f64 = area["notes"]
                        .as_array()
                        .unwrap()
                        .iter()
                        .map(|note| note["height"].as_f64().unwrap())
                        .sum();
                    assert_eq!(area["height"].as_f64().unwrap() - notes_height, 12.0);
                }
                replica
                    .doc()
                    .set_note_separator_state(opener.doc().note_separator_state().unwrap());
                let replicated: serde_json::Value = serde_json::from_str(
                    &replica.layout_document_with_regions_json(&request).unwrap(),
                )
                .unwrap();
                assert_eq!(
                    replicated["layout"]["pages"].as_array().unwrap().len(),
                    layout["pages"].as_array().unwrap().len()
                );
                assert_eq!(page_breaks(&replicated["layout"]), page_breaks(&layout));
                if kind == "footnote" {
                    assert_eq!(layout["pages"].as_array().unwrap().len(), 2);
                    assert_eq!(fallback["layout"]["pages"].as_array().unwrap().len(), 1);
                }
            }
        }
    }

    #[test]
    fn equal_separator_state_keeps_the_lowering_and_empty_state_clears_it() {
        let (opener, _) = layout_with_note_separator(
            "footnote",
            Some(
                r#"<w:p><w:pPr><w:spacing w:line="960" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="144"/></w:rPr><w:separator/></w:r></w:p>"#,
            ),
        );
        let state = opener.doc().note_separator_state().unwrap().unwrap();
        assert!(Arc::ptr_eq(
            &state,
            &opener.doc().note_separator_state().unwrap().unwrap()
        ));
        let replica = EngineSession::new(313);
        replica
            .doc()
            .set_note_separator_state(Some(Arc::clone(&state)));
        let env = RenderEnv::default();
        let lowered = replica
            .lower_note_separator("footnote", &env)
            .unwrap()
            .unwrap();
        let equal: Arc<[u8]> = Arc::from(state.as_ref());
        assert!(!Arc::ptr_eq(&state, &equal));
        replica.doc().set_note_separator_state(Some(equal));
        assert!(Arc::ptr_eq(
            &state,
            &replica.doc().note_separator_state().unwrap().unwrap()
        ));
        assert!(
            replica
                .lower_note_separator("endnote", &env)
                .unwrap()
                .is_none()
        );
        assert!(Rc::ptr_eq(
            &lowered,
            &replica
                .lower_note_separator("footnote", &env)
                .unwrap()
                .unwrap()
        ));
        replica
            .doc()
            .set_note_separator_state(Some(Arc::<[u8]>::from([])));
        assert!(replica.doc().note_separator_state().unwrap().is_none());
        assert!(
            replica
                .lower_note_separator("footnote", &env)
                .unwrap()
                .is_none()
        );
        assert!(replica.note_separators.borrow().is_none());
    }

    #[test]
    fn three_separator_paragraphs_reserve_three_lines() {
        let paragraph =
            r#"<w:p><w:pPr><w:pStyle w:val="Separator"/></w:pPr><w:r><w:separator/></w:r></w:p>"#;
        let (_, single) = layout_with_note_separator("footnote", Some(paragraph));
        let (_, triple) = layout_with_note_separator("footnote", Some(&paragraph.repeat(3)));
        let separator_reservation = |layout: &serde_json::Value| {
            let page = &layout["pages"][0];
            let area = &page["noteAreas"][0];
            page["footnoteReservedHeight"].as_f64().unwrap()
                - area["notes"][0]["height"].as_f64().unwrap()
        };
        assert!(
            (separator_reservation(&triple) - 3.0 * separator_reservation(&single)).abs() < 0.01
        );
    }

    #[test]
    fn a_separator_run_font_sizes_the_line_above_its_paragraph_default() {
        let separator = |size| {
            format!(
                r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="{size}"/></w:rPr><w:separator/></w:r></w:p>"#
            )
        };
        let (_, small) = layout_with_note_separator("footnote", Some(&separator(20)));
        let (_, large) = layout_with_note_separator("footnote", Some(&separator(72)));
        let height = |layout: &serde_json::Value| {
            layout["pages"][0]["noteAreas"][0]["separator"]["height"]
                .as_f64()
                .unwrap()
        };
        assert!(height(&large) >= 48.0);
        assert!(height(&large) > 3.0 * height(&small));
        let page = &large["pages"][0];
        assert_eq!(
            page["footnoteReservedHeight"].as_f64().unwrap(),
            height(&large) + page["noteAreas"][0]["notes"][0]["height"].as_f64().unwrap()
        );
    }

    #[test]
    fn adjacent_separator_paragraph_spacing_collapses() {
        for (contextual, expected) in [("", 72.0), ("<w:contextualSpacing/>", 42.0)] {
            let paragraphs = format!(
                r#"<w:p><w:pPr>{contextual}<w:spacing w:before="60" w:after="300" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p><w:p><w:pPr>{contextual}<w:spacing w:before="450" w:after="90" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#
            );
            let (_, layout) = layout_with_note_separator("footnote", Some(&paragraphs));
            let page = &layout["pages"][0];
            let area = &page["noteAreas"][0];
            assert_eq!(area["separator"]["height"], expected);
            assert_eq!(
                page["footnoteReservedHeight"].as_f64().unwrap(),
                expected + area["notes"][0]["height"].as_f64().unwrap()
            );
        }
    }

    #[test]
    fn separator_reservations_follow_each_sections_content_width() {
        let separator = r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r><w:r><w:t>A separator paragraph that wraps only in the narrow section.</w:t></w:r></w:p>"#;
        for kind in ["footnote", "endnote"] {
            let body = format!(
                r#"<w:p><w:r><w:t>Wide section</w:t></w:r><w:r><w:{kind}Reference w:id="1"/></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr></w:p><w:p><w:r><w:t>Narrow section</w:t></w:r><w:r><w:{kind}Reference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="4320" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#
            );
            let (engine, layout) = layout_note_separator_document(
                kind,
                Some(separator),
                &body,
                serde_json::json!([
                    {"sectionId": "wide", "properties": {
                        "pageWidth": 12240, "pageHeight": 15840,
                        "marginTop": 720, "marginRight": 720, "marginBottom": 720, "marginLeft": 720
                    }},
                    {"sectionId": "narrow", "properties": {
                        "pageWidth": 4320, "pageHeight": 15840,
                        "marginTop": 720, "marginRight": 720, "marginBottom": 720, "marginLeft": 720
                    }}
                ]),
                &[1, 2],
            );
            let pages = layout["pages"].as_array().unwrap();
            assert_eq!(pages.len(), 2);
            assert_eq!(pages[0]["sectionIndex"], 0);
            assert_eq!(pages[1]["sectionIndex"], 1);
            let pagination = engine.pagination.borrow();
            let heights = pagination
                .input
                .as_ref()
                .unwrap()
                .options
                .note_separator_heights
                .as_ref()
                .unwrap();
            let note_kind = if kind == "footnote" {
                NoteKind::Footnote
            } else {
                NoteKind::Endnote
            };
            assert_eq!(heights.height(note_kind, 0), 16.0);
            assert!(heights.height(note_kind, 1) > heights.height(note_kind, 0));
            if kind == "endnote" {
                assert!(pages[0].get("noteAreas").is_none());
                assert_eq!(
                    pages[1]["noteAreas"][0]["notes"].as_array().unwrap().len(),
                    2
                );
            }
            for page in pages {
                let Some(areas) = page["noteAreas"].as_array() else {
                    continue;
                };
                let area = &areas[0];
                let section = page["sectionIndex"].as_u64().unwrap() as usize;
                let separator_height = heights.height(note_kind, section);
                assert_eq!(area["separator"]["height"], separator_height);
                assert!(area["separator"].get("blocks").is_none());
                let note_height: f64 = area["notes"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|note| note["height"].as_f64().unwrap())
                    .sum();
                assert_eq!(
                    page["footnoteReservedHeight"].as_f64().unwrap(),
                    separator_height + note_height
                );
            }
        }
    }

    #[test]
    fn note_text_is_painted_below_the_measured_separator() {
        let text_offset = |separator: Option<&str>| {
            let (engine, layout) = layout_with_note_separator("footnote", separator);
            engine.build_display_list_frame("{}", 0).unwrap();
            let baseline = engine
                .with_display_list(|list| {
                    list.pages[0].note_areas[0]
                        .primitives
                        .iter()
                        .find_map(|primitive| match primitive {
                            docx_layout::display_list::Primitive::Text(run) => {
                                run.baseline_y.as_f64()
                            }
                            docx_layout::display_list::Primitive::GlyphRun(run) => {
                                run.glyphs.first().map(|glyph| glyph.y)
                            }
                            _ => None,
                        })
                })
                .flatten()
                .unwrap();
            baseline - layout["pages"][0]["noteAreas"][0]["y"].as_f64().unwrap()
        };
        let measured = text_offset(Some(
            r#"<w:p><w:pPr><w:spacing w:line="960" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#,
        ));
        assert!((measured - text_offset(None) - (64.0 - 12.0)).abs() < 0.01);
    }

    #[test]
    fn a_missing_separator_keeps_twelve_pixels_of_reservation() {
        let (engine, layout) = layout_with_note_separator("footnote", None);
        assert_eq!(layout["pages"][0]["footnoteReservedHeight"], 16.0 + 12.0);
        assert!(
            layout["pages"][0]["noteAreas"][0]
                .get("separator")
                .is_none()
        );
        let pagination = engine.pagination.borrow();
        let options = serde_json::to_value(&pagination.input.as_ref().unwrap().options).unwrap();
        assert!(options.get("noteSeparatorHeights").is_none());
    }

    #[test]
    fn a_separator_without_paragraphs_keeps_twelve_pixels_of_reservation() {
        let (_, layout) = layout_with_note_separator("footnote", Some(""));
        assert_eq!(layout["pages"][0]["footnoteReservedHeight"], 16.0 + 12.0);
        assert!(
            layout["pages"][0]["noteAreas"][0]
                .get("separator")
                .is_none()
        );
    }

    #[test]
    fn endnotes_reserve_their_own_separator_height() {
        let (_, layout) = layout_with_note_separator(
            "endnote",
            Some(
                r#"<w:p><w:pPr><w:spacing w:line="960" w:lineRule="exact"/></w:pPr><w:r><w:separator/></w:r></w:p>"#,
            ),
        );
        let area = layout["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["noteAreas"].as_array().into_iter().flatten())
            .find(|area| area["kind"] == "endnote")
            .unwrap();
        assert_eq!(area["height"], 64.0 + 16.0);
    }

    #[test]
    fn a_preview_seeded_from_a_body_prefix_paints_the_first_pages_of_the_document() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let styles = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>"#;
        let cell = |text: &str| format!("<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>");
        let mut body = String::new();
        for index in 0..160 {
            if index % 12 == 0 {
                body.push_str(&format!(
                    r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Part {index}</w:t></w:r></w:p>"#
                ));
            }
            if index % 40 == 5 {
                body.push_str(&format!(
                    "<w:tbl><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl>",
                    cell("North"),
                    cell("South"),
                    cell(&format!("Row {index}")),
                    cell("A cell long enough to wrap onto a second line of its column"),
                ));
            }
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index}: a sentence that wraps across the narrow column of this small page, twice over.</w:t></w:r></w:p>"
            ));
            if index == 150 {
                body.push_str(r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="5760" w:h="4320"/></w:sectPr></w:pPr></w:p>"#);
            }
        }
        let bytes = docx_bytes(styles, &body);
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [
                { "sectionId": "first", "properties": {
                    "pageWidth": 5760, "pageHeight": 4320,
                    "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
                } },
                { "sectionId": "last", "properties": {
                    "pageWidth": 12240, "pageHeight": 2880,
                    "marginTop": 200, "marginRight": 200, "marginBottom": 200, "marginLeft": 200
                } }
            ] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let extras =
            serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
        let first_pages = |engine: &EngineSession, layout_json: String| {
            engine.build_display_list_frame(&extras, 0).unwrap();
            let pages = engine
                .with_display_list(|list| list.pages[..3].to_vec())
                .unwrap();
            (layout_json, pages)
        };

        let full = EngineSession::new(310);
        crate::seed::seed_from_docx(full.doc(), &bytes).unwrap();
        let (_, full_pages) = first_pages(
            &full,
            full.layout_document_with_regions_retained_json(&request)
                .unwrap(),
        );
        let preview = EngineSession::new(311);
        assert!(crate::seed::seed_docx_preview(preview.doc(), &bytes, 60).unwrap());
        assert!(
            preview.doc().paragraphs("body").unwrap().len()
                < full.doc().paragraphs("body").unwrap().len()
        );
        let cell_stories = |engine: &EngineSession| {
            use yrs::{Map, ReadTxn, Transact};
            let txn = engine.doc().yrs_doc().transact();
            let stories = txn.get_map(crate::STORIES).unwrap();
            stories
                .iter(&txn)
                .filter(|(id, _)| id.starts_with("body:"))
                .count()
        };
        // Only the tables the preview reaches keep their cell stories.
        assert!(cell_stories(&preview) > 0 && cell_stories(&preview) < cell_stories(&full));
        // The preview ends mid-section, so only a prefix pass that stops
        // short of its end lays out pages the rest cannot move.
        let (layout, preview_pages) = first_pages(
            &preview,
            preview
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap(),
        );
        assert!(layout.contains("\"provisional\":true"));
        assert_eq!(preview_pages, full_pages);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn a_preview_draws_the_images_its_first_pages_use() {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        // A 1x1 PNG, and a second medium nothing on the first pages uses.
        const PNG: &[u8] = &[
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52,
            0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0d, 0x49,
            0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0xf0, 0x1f, 0, 0x05, 0, 0x01,
            0xff, 0x7d, 0x31, 0x96, 0xd5, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60,
            0x82,
        ];
        let image = |id: &str| {
            format!(
                r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="{id}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
            )
        };
        let mut body = image("rIdFirst");
        for index in 0..150 {
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index} of a document whose first page shows a picture.</w:t></w:r></w:p>"
            ));
        }
        body.push_str(&image("rIdLater"));
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>{body}</w:body></w:document>"#
        );
        let bytes = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), PACKAGE_RELS.as_bytes().to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/first.png"/><Relationship Id="rIdLater" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/later.png"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), document.into_bytes()),
            ("word/media/first.png".to_owned(), PNG.to_vec()),
            ("word/media/later.png".to_owned(), PNG.to_vec()),
        ])
        .unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [{ "sectionId": "main", "properties": {
                "pageWidth": 5760, "pageHeight": 4320,
                "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
            } }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": { "mediaTokens": true }
        })
        .to_string();
        let extras =
            serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
        let first_page = |engine: &EngineSession| {
            engine.build_display_list_frame(&extras, 0).unwrap();
            engine
                .with_display_list(|list| list.pages[0].clone())
                .unwrap()
        };
        let full = EngineSession::new(312);
        crate::seed::seed_with_layout_tokens(full.doc(), &bytes).unwrap();
        full.layout_document_with_regions_retained_json(&request)
            .unwrap();
        let preview = EngineSession::new(312);
        assert!(crate::seed::seed_docx_preview(preview.doc(), &bytes, 40).unwrap());
        preview
            .layout_document_with_regions_prefix_retained_json(&request, 3)
            .unwrap();
        let page = first_page(&preview);
        let painted = serde_json::to_string(&page).unwrap();
        assert!(painted.contains(r#""media:0""#) && !painted.contains("data:"));
        assert_eq!(page, first_page(&full));
        assert_eq!(
            preview.doc().media_table().unwrap().bytes(0).unwrap(),
            full.doc().media_table().unwrap().bytes(0).unwrap()
        );
        docx_layout::clear_measure_fonts();
    }

    /// Times opening a document for display in full and from a body prefix.
    /// Run: `DOCX_PREVIEW_PROBE=<file.docx> cargo test -p betteroffice-docx-edit --release --lib -- --ignored preview_probe --nocapture`
    #[test]
    #[ignore = "perf probe; run explicitly with DOCX_PREVIEW_PROBE set"]
    fn preview_probe() {
        let Ok(path) = std::env::var("DOCX_PREVIEW_PROBE") else {
            return;
        };
        let blocks = std::env::var("DOCX_PREVIEW_BLOCKS")
            .ok()
            .and_then(|value| value.parse().ok())
            .unwrap_or(200);
        let bytes = std::fs::read(path).unwrap();
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": { "sections": [{ "sectionId": "main", "properties": {} }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let extras =
            serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
        let open = |engine: &EngineSession, seed: &dyn Fn(&EngineSession)| {
            let started = std::time::Instant::now();
            seed(engine);
            let seeded = started.elapsed();
            engine
                .layout_document_with_regions_prefix_retained_json(&request, 3)
                .unwrap();
            let laid_out = started.elapsed();
            engine.build_display_list_frame(&extras, 0).unwrap();
            let painted = started.elapsed();
            let pages = engine
                .pagination
                .borrow()
                .layout
                .as_ref()
                .unwrap()
                .pages
                .len();
            println!("  seed {seeded:?}, +layout {laid_out:?}, +frame {painted:?}, {pages} pages");
            engine
                .with_display_list(|list| list.pages[..3].to_vec())
                .unwrap()
        };
        println!("full open:");
        let full = open(&EngineSession::new(320), &|engine| {
            crate::seed::seed_from_docx(engine.doc(), &bytes).unwrap();
        });
        println!("preview of the first {blocks} body blocks:");
        let preview = open(&EngineSession::new(320), &|engine| {
            assert!(crate::seed::seed_docx_preview(engine.doc(), &bytes, blocks).unwrap());
        });
        for (index, (full, preview)) in full.iter().zip(&preview).enumerate() {
            println!("  page {index} identical: {}", full == preview);
        }
    }

    /// Three sections of ten paragraphs under a header taller than the top
    /// margin, and the region request that lays them out.
    fn sectioned_header_fixture() -> (Vec<u8>, String) {
        let section = r#"<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="100" w:footer="100"/></w:sectPr>"#;
        let mut body = String::new();
        for index in 0..30 {
            let properties = if index % 10 == 9 { section } else { "" };
            body.push_str(&format!(
                r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>{properties}</w:pPr><w:r><w:t>Paragraph {index} of the body</w:t></w:r></w:p>"#
            ));
        }
        body.push_str(section);
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let sections: Vec<_> = (0..3)
            .map(|index| {
                serde_json::json!({
                    "sectionId": format!("s{index}"),
                    "pageSize": {"w": 300, "h": 200},
                    "margins": {"top": 20, "right": 20, "bottom": 20, "left": 20, "header": 5, "footer": 5},
                    "headerFooterRefs": {"headerDefault": "rId1"}
                })
            })
            .collect();
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": sections},
            "measurement": {
                "fontChains": {"liberation sans|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        (docx_bytes("", &body), request.to_string())
    }

    fn open_sectioned(bytes: &[u8], client_id: u64) -> EngineSession {
        let engine = EngineSession::new(client_id);
        crate::seed::seed_from_docx(engine.doc(), bytes).unwrap();
        engine
            .doc()
            .create_story(
                "hf:rId1",
                "A header taller than the top margin",
                "Normal",
                "left",
            )
            .unwrap();
        engine
    }

    fn insert_x(engine: &EngineSession, story: &str, at: u32) {
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new(story, at),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
    }

    fn retained_pages(engine: &EngineSession) -> serde_json::Value {
        serde_json::to_value(&engine.pagination.borrow().layout.as_ref().unwrap().pages).unwrap()
    }

    fn assert_fingerprints_as_paginated(engine: &EngineSession) {
        let pagination = engine.pagination.borrow();
        assert_eq!(
            pagination.block_fingerprints,
            measured_fingerprints(pagination.input.as_ref().unwrap()).unwrap(),
            "retained fingerprints are those of the arena as paginated"
        );
    }

    #[test]
    fn the_first_edit_after_a_region_open_reuses_section_breaks_and_converges() {
        let (bytes, request) = sectioned_header_fixture();
        let engine = open_sectioned(&bytes, 142);
        let output: serde_json::Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        assert!(
            output["options"]["margins"]["top"].as_f64().unwrap() > 20.0,
            "the header widens the margins"
        );
        assert!(output["layout"]["pages"].as_array().unwrap().len() > 3);
        engine.build_display_list_frame("{}", 0).unwrap();

        insert_x(&engine, "body", 2);
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let after = engine.stats();
        assert_eq!(
            after.resident_measure_calls - before.resident_measure_calls,
            1,
            "only the edited paragraph re-measures; the section breaks keep their extents"
        );
        assert_eq!(
            after.incremental_pagination_calls - before.incremental_pagination_calls,
            1
        );
        assert_eq!(
            after.rebuilt_pages, 1,
            "pagination converges after the edited page"
        );
        assert_fingerprints_as_paginated(&engine);

        let reference = open_sectioned(&bytes, 143);
        insert_x(&reference, "body", 2);
        reference
            .layout_document_with_regions_json(&request)
            .unwrap();
        assert_eq!(retained_pages(&engine), retained_pages(&reference));
    }

    #[test]
    fn the_first_edit_after_an_open_builds_on_the_hosts_frame() {
        fn integral(value: &mut serde_json::Value) {
            match value {
                serde_json::Value::Number(number) => {
                    if let Some(whole) = number.as_f64().filter(|value| value.fract() == 0.0) {
                        *value = serde_json::json!(whole as i64);
                    }
                }
                serde_json::Value::Array(items) => items.iter_mut().for_each(integral),
                serde_json::Value::Object(fields) => fields.values_mut().for_each(integral),
                _ => {}
            }
        }
        let (bytes, request) = sectioned_header_fixture();
        let open = |client_id| {
            let engine = open_sectioned(&bytes, client_id);
            let output: serde_json::Value =
                serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                    .unwrap();
            // A host writes whole numbers without a fraction.
            let mut headers_footers = output["headersFooters"].clone();
            assert!(headers_footers.is_object());
            integral(&mut headers_footers);
            let extras = serde_json::json!({ "headersFooters": headers_footers }).to_string();
            engine.build_display_list_frame(&extras, 0).unwrap();
            (engine, extras)
        };
        let (engine, _) = open(146);
        insert_x(&engine, "body", 2);
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds - before.incremental_display_builds,
            1
        );

        let (reference, extras) = open(146);
        insert_x(&reference, "body", 2);
        reference
            .layout_document_with_regions_json(&request)
            .unwrap();
        reference.build_display_list_frame(&extras, 0).unwrap();
        let pages = |engine: &EngineSession| engine.with_display_list(|list| list.pages.clone());
        assert_eq!(pages(&engine), pages(&reference));
    }

    #[test]
    fn a_taller_header_refingerprints_the_reused_section_breaks() {
        let (bytes, request) = sectioned_header_fixture();
        let engine = open_sectioned(&bytes, 144);
        engine.layout_document_with_regions_json(&request).unwrap();
        let grow = |engine: &EngineSession| {
            let ctx = crate::EditCtx::local("", "");
            for _ in 0..3 {
                engine
                    .doc()
                    .split_paragraph(&ctx, crate::Position::new("hf:rId1", 1), None)
                    .unwrap();
            }
        };
        grow(&engine);
        engine.layout_document_with_regions_json(&request).unwrap();
        assert_fingerprints_as_paginated(&engine);

        let reference = open_sectioned(&bytes, 145);
        grow(&reference);
        reference
            .layout_document_with_regions_json(&request)
            .unwrap();
        assert_eq!(retained_pages(&engine), retained_pages(&reference));
    }

    #[test]
    fn a_middle_section_whose_header_grows_repaginates_from_its_start() {
        let (bytes, request) = sectioned_header_fixture();
        let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
        for (index, section) in request["regions"]["sections"]
            .as_array_mut()
            .unwrap()
            .iter_mut()
            .enumerate()
        {
            section["headerFooterRefs"]["headerDefault"] = format!("rId{}", index + 1).into();
        }
        let request = request.to_string();
        let open = |client_id| {
            let engine = open_sectioned(&bytes, client_id);
            for story in ["hf:rId2", "hf:rId3"] {
                engine
                    .doc()
                    .create_story(story, "A header", "Normal", "left")
                    .unwrap();
            }
            engine
        };
        let grow = |engine: &EngineSession| {
            let ctx = crate::EditCtx::local("", "");
            for _ in 0..3 {
                engine
                    .doc()
                    .split_paragraph(&ctx, crate::Position::new("hf:rId2", 1), None)
                    .unwrap();
            }
        };
        let engine = open(146);
        engine.layout_document_with_regions_json(&request).unwrap();
        grow(&engine);
        engine.layout_document_with_regions_json(&request).unwrap();

        let reference = open(147);
        grow(&reference);
        reference
            .layout_document_with_regions_json(&request)
            .unwrap();
        assert_eq!(retained_pages(&engine), retained_pages(&reference));
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

    /// An engine laid out over small pages: an editable first paragraph, then `fillers`
    /// one-line paragraphs. Returns it with the display extras.
    fn paged_filler_engine(client_id: u64, fillers: usize) -> (EngineSession, String) {
        paged_engine(
            client_id,
            &format!(
                "<w:p><w:r><w:t>Editable paragraph</w:t></w:r></w:p>{}",
                "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(fillers)
            ),
        )
    }

    fn paged_engine(client_id: u64, body: &str) -> (EngineSession, String) {
        let (engine, extras, output) = paged_region_engine(client_id, body);
        let font_chains =
            serde_json::from_str::<serde_json::Value>(&extras).unwrap()["fontChains"].clone();
        for measured_block in output["measured"].as_array().unwrap() {
            if measured_block["block"]["kind"] != "paragraph" {
                continue;
            }
            let template = serde_json::json!({
                "block": measured_block["block"],
                "maxWidth": 248,
                "fontChains": font_chains,
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
        (
            engine,
            serde_json::json!({ "fontChains": font_chains }).to_string(),
        )
    }

    /// [`paged_engine`] with its region state kept, so an edit in a table cell
    /// relays out through the region pass as it does in the worker.
    fn paged_region_engine(
        client_id: u64,
        body: &str,
    ) -> (EngineSession, String, serde_json::Value) {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let engine = EngineSession::new(client_id);
        crate::seed::seed_from_docx(engine.doc(), &docx_bytes("", body)).unwrap();
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
        // As the host's extras carry the pass's headers/footers, so region
        // edits keep building on the shown frame.
        let mut extras = serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } });
        if let Some(headers_footers) = output.get("headersFooters") {
            extras["headersFooters"] = headers_footers.clone();
        }
        (engine, extras.to_string(), output)
    }

    fn full_display_build(
        engine: &EngineSession,
        extras: &str,
    ) -> docx_layout::display_list::DisplayList {
        let pagination = engine.pagination.borrow();
        docx_layout::build_display_list_value_from_resident(
            pagination.input.as_ref().unwrap(),
            pagination.layout.as_ref().unwrap(),
            extras,
        )
        .unwrap()
    }

    /// Collapsed backspaces retain cached paragraphs and exact windowed frames.
    #[test]
    fn windowed_backspaces_keep_the_paragraph_index_and_match_a_full_build() {
        use std::sync::Arc;
        use yrs::{Assoc, IndexedSequence};

        docx_layout::clear_measure_fonts();
        let font = docx_layout::register_measure_font(LIBERATION).unwrap();
        let engine = EngineSession::new(219);
        engine.set_local_lowering(true);
        crate::seed::seed_from_docx(
            engine.doc(),
            &docx_bytes(
                "",
                &"<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(49),
            ),
        )
        .unwrap();
        let paragraph = engine.doc().paragraphs("body").unwrap().remove(24);
        let start = engine
            .doc()
            .paragraph_index("body")
            .unwrap()
            .para_span(&paragraph.para_id)
            .unwrap()
            .0;
        let at = start + 7;
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", at),
                "😀x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        engine
            .layout_document_with_regions_json(&small_page_request(font))
            .unwrap();
        let initial = engine.build_display_list_frame("{}", 0).unwrap();
        let caret_page = {
            let pagination = engine.pagination.borrow();
            let layout = pagination.layout.as_ref().unwrap();
            assert!(layout.pages.len() >= 4);
            layout
                .pages
                .iter()
                .position(|page| {
                    page.fragments.iter().any(|fragment| {
                        matches!(fragment, Fragment::Paragraph(value)
                            if block_key(&value.block_id) == paragraph.para_id)
                    })
                })
                .unwrap()
        };
        let txn = engine.doc().yrs_doc().transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        let head = story.sticky_index(&txn, at + 3, Assoc::After).unwrap();
        drop(txn);
        engine.set_resident_caret_head(Some(("body".to_owned(), head)));
        engine.set_display_window(Some(caret_page..caret_page + 1));
        engine.set_windowed_incremental_builds(true);
        engine.doc().paragraph_index("body").unwrap();
        engine.doc().segment_index("body").unwrap();
        let mut retained = HashMap::new();
        let initial = crate::frame_delta::apply_placeholder_test_frame(&initial, &mut retained);
        assert_eq!(initial, full_display_build(&engine, "{}"));

        for units in [1, 2] {
            let txn = engine.doc().yrs_doc().transact();
            let end = engine
                .resident_caret_head
                .borrow()
                .as_ref()
                .unwrap()
                .1
                .get_offset(&txn)
                .unwrap()
                .index;
            drop(txn);
            let before = engine.doc().committed_epoch();
            engine
                .edit_resident_text(crate::StoryRange::new("body", end - units, end), None, true)
                .unwrap();
            let after = engine.doc().committed_epoch();
            assert_eq!(after, before + 1);
            let shifted = engine
                .doc()
                .paragraph_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .expect("the delete advances the cached paragraph index");
            let shifted_segments = engine
                .doc()
                .segment_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .expect("the delete advances the cached segment index");
            let epoch = engine.display.borrow().binary_frame_epoch;
            let frame = engine.apply_and_layout("body", epoch).unwrap();
            assert!(engine.pagination.borrow().last_incremental);
            assert!(Arc::ptr_eq(
                &shifted,
                &engine.doc().paragraph_index("body").unwrap(),
            ));
            assert!(Arc::ptr_eq(
                &shifted_segments,
                &engine.doc().segment_index("body").unwrap(),
            ));
            let expected = full_display_build(&engine, "{}");
            assert_eq!(engine.with_display_list(Clone::clone).unwrap(), expected);
            assert_eq!(
                crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained),
                expected,
            );
        }
        docx_layout::clear_measure_fonts();
    }

    /// Embed removals and paragraph merges leave the story caches stale.
    #[test]
    fn resident_structural_deletes_do_not_advance_the_story_indexes() {
        for (kind, start, end) in [("sdt", 1, 2), ("pilcrow", 1, 2), ("pageBreak", 0, 1)] {
            let engine = EngineSession::new(222);
            engine.set_local_lowering(true);
            let ctx = crate::EditCtx::local("", "");
            engine
                .doc()
                .create_story("body", "ABC", "Normal", "left")
                .unwrap();
            if kind == "pilcrow" {
                engine
                    .doc()
                    .split_paragraph(&ctx, crate::Position::new("body", 1), None)
                    .unwrap();
            } else {
                engine
                    .doc()
                    .apply_raw_ops(
                        "body",
                        vec![crate::RawOp::InsertEmbed {
                            index: 1,
                            kind: kind.to_owned(),
                            payload: vec![("embedId".to_owned(), yrs::Any::from("control-1"))],
                            attrs: Default::default(),
                        }],
                        &ctx,
                    )
                    .unwrap();
            }
            engine.doc().paragraph_index("body").unwrap();
            engine.doc().segment_index("body").unwrap();
            let before = engine.doc().committed_epoch();
            engine
                .edit_resident_text(crate::StoryRange::new("body", start, end), None, true)
                .unwrap();
            let after = engine.doc().committed_epoch();
            assert_eq!(after, before + 1);
            assert!(
                engine
                    .doc()
                    .paragraph_indexes
                    .lock()
                    .unwrap()
                    .get("body", after)
                    .is_none()
            );
            assert!(
                engine
                    .doc()
                    .segment_indexes
                    .lock()
                    .unwrap()
                    .get("body", after)
                    .is_none()
            );
        }
    }

    #[test]
    fn resident_random_text_edits_keep_story_indexes_exact() {
        use crate::segments::tests::{
            Unit, assert_paragraph_indexes_eq, assert_segment_invariants, next_random, random_text,
            seed_text_stream, units,
        };

        for seed in [1, 7, 42, 0x1234_5678] {
            let engine = EngineSession::new(224);
            engine.set_local_lowering(true);
            seed_text_stream(engine.doc());
            let mut random = seed;
            let mut steps = 0;
            let mut plain_edits = 0;
            let mut advanced_both = 0;
            let mut deleted_pilcrow = false;
            let mut deleted_embed = false;
            let mut deleted_surrogate_pair = false;
            for step in 0..200 {
                let doc = engine.doc();
                let segments = doc.segment_index("body").unwrap();
                doc.paragraph_index("body").unwrap();
                let before_units = units(&segments);
                let before = doc.committed_epoch();
                let len = doc.story_len("body").unwrap();
                let (start, end, text) = if step < 3 {
                    let start = before_units
                        .iter()
                        .position(|unit| match step {
                            0 => matches!(unit, Unit::Text(value) if (0xd800..=0xdbff).contains(value)),
                            1 => matches!(unit, Unit::Embed),
                            _ => matches!(unit, Unit::Pilcrow),
                        })
                        .unwrap() as u32;
                    (start, start + 1 + u32::from(step == 0), None)
                } else if next_random(&mut random) % 3 != 0 {
                    let start = next_random(&mut random) % (len + 1);
                    let start = start - u32::from(!segments.is_char_boundary(start));
                    (start, start, Some(random_text(&mut random)))
                } else {
                    let start = next_random(&mut random) % len;
                    let start = start - u32::from(!segments.is_char_boundary(start));
                    let end = (start + 1 + next_random(&mut random) % 3).min(len);
                    let end = end + u32::from(!segments.is_char_boundary(end));
                    (start, end, None)
                };
                if engine
                    .edit_resident_text(crate::StoryRange::new("body", start, end), text, true)
                    .is_err()
                {
                    continue;
                }
                if text.is_none() {
                    let removed = &before_units[start as usize..end as usize];
                    deleted_pilcrow |= removed.iter().any(|unit| matches!(unit, Unit::Pilcrow));
                    deleted_embed |= removed.iter().any(|unit| matches!(unit, Unit::Embed));
                    deleted_surrogate_pair |= step == 0;
                }
                let after = doc.committed_epoch();
                let after_len = doc.story_len("body").unwrap();
                let plain = match text {
                    Some(text) => {
                        after_len.checked_sub(len) == Some(text.encode_utf16().count() as u32)
                    }
                    None => {
                        segments.is_text_range(start, end)
                            && len.checked_sub(after_len) == Some(end - start)
                    }
                };
                let cached_segments = doc.segment_indexes.lock().unwrap().get("body", after);
                let cached_paragraphs = doc.paragraph_indexes.lock().unwrap().get("body", after);
                if after != before {
                    steps += 1;
                    if plain {
                        plain_edits += 1;
                        advanced_both +=
                            usize::from(cached_segments.is_some() && cached_paragraphs.is_some());
                    }
                }
                let txn = doc.yrs_doc().transact();
                let story = crate::story_ref(&txn, "body").unwrap();
                let (cold_segments, cold_paragraphs) = crate::segments::build_indexes(&story, &txn);
                drop(txn);
                assert_segment_invariants(&cold_segments);
                if let Some(cached) = cached_segments {
                    assert_eq!(
                        units(&cached),
                        units(&cold_segments),
                        "seed {seed}, step {step}"
                    );
                    assert_segment_invariants(&cached);
                }
                if let Some(cached) = cached_paragraphs {
                    assert_paragraph_indexes_eq(&cached, &cold_paragraphs);
                }
                doc.segment_index("body").unwrap();
                doc.paragraph_index("body").unwrap();
            }
            assert!(deleted_pilcrow && deleted_embed && deleted_surrogate_pair);
            assert!(steps > 150, "seed {seed}: {steps} committed edits");
            assert!(
                advanced_both * 2 > plain_edits,
                "seed {seed}: {advanced_both}/{plain_edits} plain edits advanced both indexes"
            );
            assert!(
                advanced_both * 5 > steps * 2,
                "seed {seed}: {advanced_both}/{steps} edits advanced both indexes"
            );
        }
    }

    /// A delete that splits a surrogate pair keeps the paragraph index exact.
    #[test]
    fn a_delete_inside_a_surrogate_pair_keeps_the_paragraph_index_exact() {
        let engine = EngineSession::new(223);
        engine.set_local_lowering(true);
        let ctx = crate::EditCtx::local("", "");
        engine
            .doc()
            .create_story("body", "A\u{1F600}BC", "Normal", "left")
            .unwrap();
        engine
            .doc()
            .split_paragraph(&ctx, crate::Position::new("body", 4), None)
            .unwrap();
        engine.doc().paragraph_index("body").unwrap();
        engine
            .edit_resident_text(crate::StoryRange::new("body", 1, 2), None, true)
            .unwrap();
        let after = engine.doc().committed_epoch();
        let cached = engine
            .doc()
            .paragraph_indexes
            .lock()
            .unwrap()
            .take("body", after);
        let fresh = engine.doc().paragraph_index("body").unwrap();
        if let Some(cached) = cached {
            let len = engine.doc().story_len("body").unwrap();
            for index in 0..=len {
                let geometry = |entry: Option<&crate::segments::ParaEntry>| {
                    entry.map(|entry| {
                        (
                            entry.para_id.clone(),
                            entry.start,
                            entry.pilcrow,
                            entry.node_start,
                        )
                    })
                };
                assert_eq!(
                    geometry(cached.para_at(index)),
                    geometry(fresh.para_at(index)),
                    "index {index}"
                );
            }
        }
    }

    /// Windowed edits shift a suffix once and later page builds encode only requested pages.
    #[test]
    fn windowed_edit_range_shifts_decode_to_full_build_and_keep_page_builds_scoped() {
        let (engine, extras) = paged_filler_engine(218, 48);
        let full = engine.build_display_list_frame(&extras, 0).unwrap();
        let mut retained = HashMap::new();
        let initial = crate::frame_delta::apply_placeholder_test_frame(&full, &mut retained);
        assert!(initial.pages.len() >= 4);
        assert_eq!(initial, full_display_build(&engine, &extras));
        let next_page_id = engine.display.borrow().next_page_id;
        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 3),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let bytes = engine.apply_and_layout("body", epoch).unwrap();
        let operation_count = u32::from_le_bytes(bytes[52..56].try_into().unwrap()) as usize;
        let ranges: Vec<_> = (0..operation_count)
            .map(|op| crate::frame_delta::FRAME_HEADER_LEN + op * crate::frame_delta::PAGE_OP_LEN)
            .filter(|&record| bytes[record] == crate::frame_delta::PAGE_OP_SHIFT_RANGE)
            .collect();
        assert_eq!(ranges.len(), 1);
        let record = ranges[0];
        let start = u32::from_le_bytes(bytes[record + 4..record + 8].try_into().unwrap()) as usize;
        let count =
            u32::from_le_bytes(bytes[record + 24..record + 28].try_into().unwrap()) as usize;
        assert!(start > 0 && start < initial.pages.len());
        assert_eq!(start + count, initial.pages.len());
        assert_eq!(
            i64::from_le_bytes(bytes[record + 32..record + 40].try_into().unwrap()),
            1
        );
        let expected = full_display_build(&engine, &extras);
        assert_eq!(
            crate::frame_delta::apply_placeholder_test_frame(&bytes, &mut retained),
            expected
        );
        assert_eq!(engine.with_display_list(Clone::clone).unwrap(), expected);
        assert_eq!(engine.display.borrow().next_page_id, next_page_id);

        let last = initial.pages.len() - 1;
        let epoch = engine.display.borrow().binary_frame_epoch;
        let release = engine.release_display_pages_frame(&[last], epoch).unwrap();
        crate::frame_delta::apply_placeholder_test_frame(&release, &mut retained);
        let mut before_build = engine.display.borrow().pages.clone();
        for snapshot in &mut before_build {
            snapshot.materialize_positions();
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        let built = engine
            .build_display_pages_frame(&[last, last, 0], epoch)
            .unwrap();
        assert_eq!(u32::from_le_bytes(built[52..56].try_into().unwrap()), 1);
        assert_eq!(
            built[crate::frame_delta::FRAME_HEADER_LEN],
            crate::frame_delta::PAGE_OP_UPSERT
        );
        assert_eq!(
            u32::from_le_bytes(
                built[crate::frame_delta::FRAME_HEADER_LEN + 4
                    ..crate::frame_delta::FRAME_HEADER_LEN + 8]
                    .try_into()
                    .unwrap()
            ) as usize,
            last
        );
        assert_eq!(
            crate::frame_delta::apply_placeholder_test_frame(&built, &mut retained),
            expected
        );
        let mut after_build = engine.display.borrow().pages.clone();
        for snapshot in &mut after_build {
            snapshot.materialize_positions();
        }
        assert_eq!(after_build[..last], before_build[..last]);
        assert_eq!(engine.display.borrow().next_page_id, next_page_id);

        engine.reset_frame_base();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let recovery = engine.build_display_pages_frame(&[], epoch).unwrap();
        assert_eq!(
            u32::from_le_bytes(recovery[12..16].try_into().unwrap()),
            crate::frame_delta::FRAME_FLAG_FULL
        );
        assert_eq!(
            crate::frame_delta::apply_placeholder_test_frame(&recovery, &mut retained),
            expected
        );
        docx_layout::clear_measure_fonts();
    }

    fn selective_display_engine() -> (EngineSession, serde_json::Value, String) {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let section = r#"<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400"/></w:sectPr>"#;
        let rows: String = (0..24).map(|index| format!(
            r#"<w:tr><w:trPr>{}</w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell {index}</w:t></w:r></w:p></w:tc></w:tr>"#,
            if index == 0 { "<w:tblHeader/>" } else { "" }
        )).collect();
        let mut body = format!(
            r#"<w:p><w:r><w:t>Editable paragraph</w:t></w:r><w:r><w:footnoteReference w:id="5"/></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>{rows}</w:tbl>"#
        );
        for index in 0..30 {
            body.push_str(&format!(
                r#"<w:p><w:pPr>{}</w:pPr><w:r><w:t>Filler paragraph {index}</w:t></w:r></w:p>"#,
                if index % 10 == 9 { section } else { "" }
            ));
        }
        body.push_str(section);
        let engine = EngineSession::new(217);
        crate::seed::seed_from_docx(engine.doc(), &docx_bytes("", &body)).unwrap();
        let fields = [serde_json::json!({
            "type": "paragraph", "content": [
                {"type": "simpleField", "fieldType": "PAGE", "instruction": " PAGE ",
                    "content": [{"type": "run", "content": [{"type": "text", "text": "1"}]}]},
                {"type": "simpleField", "fieldType": "NUMPAGES", "instruction": " NUMPAGES ",
                    "content": [{"type": "run", "content": [{"type": "text", "text": "97"}]}]}
            ]
        })];
        crate::seed::seed_blocks(
            engine.doc(),
            None,
            &[
                ("hf:header".to_owned(), &fields),
                ("hf:footer".to_owned(), &fields),
            ],
        )
        .unwrap();
        engine
            .doc()
            .create_story("fn:5", "Footnote text", "Normal", "left")
            .unwrap();
        let shape = serde_json::json!({
            "shapeType": "rect", "size": {"width": 190500, "height": 190500},
            "position": {
                "horizontal": {"relativeTo": "column", "posOffset": 0},
                "vertical": {"relativeTo": "paragraph", "posOffset": 0}
            },
            "wrap": {"type": "inFront"}
        });
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![crate::RawOp::InsertEmbed {
                    index: 1,
                    kind: "shape".to_owned(),
                    payload: vec![("shapeJson".to_owned(), Any::from(shape.to_string()))],
                    attrs: Attrs::new(),
                }],
                &crate::EditCtx::local("", ""),
            )
            .unwrap();
        let deletion = engine
            .doc()
            .delete_range(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::StoryRange::new("body", 3, 4),
            )
            .unwrap();
        let sections: Vec<_> = (0..3)
            .map(|index| {
                serde_json::json!({
                    "sectionId": format!("s{index}"),
                    "pageSize": {"w": 300 + index * 30, "h": 200 + index * 20},
                    "margins": {"top": 20, "right": 20, "bottom": 20, "left": 20,
                        "header": 5, "footer": 5},
                    "headerFooterRefs": {"headerDefault": "header", "footerDefault": "footer"},
                    "pageNumbering": {"start": 1,
                        "format": if index == 0 { "lowerRoman" } else { "decimal" }}
                })
            })
            .collect();
        let request = serde_json::json!({
            "bodyStory": "body", "regions": {"sections": sections},
            "notes": {"contents": [{"id": 5, "noteKind": "footnote", "height": 0}]},
            "measurement": {
                "fontChains": {"liberation sans|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        (engine, request, deletion.revision_ids[0].clone())
    }

    #[test]
    fn selective_display_conversion_matches_legacy_after_preview_edit_and_release() {
        use docx_layout::display_list::resident_conversion_test_support::with_layout_conversion;

        let mut runs = Vec::new();
        for force_full in [false, true] {
            let (engine, mut request, revision_id) = selective_display_engine();
            let font_chains = request["measurement"]["fontChains"].clone();
            let extras = |output: &serde_json::Value| {
                serde_json::json!({
                    "fontChains": font_chains,
                    "headersFooters": output["headersFooters"]
                })
                .to_string()
            };
            let output: serde_json::Value = serde_json::from_str(
                &engine
                    .layout_document_with_regions_json(&request.to_string())
                    .unwrap(),
            )
            .unwrap();
            engine
                .build_display_list_frame(&extras(&output), 0)
                .unwrap();
            engine.set_display_window(Some(0..1));
            engine.set_windowed_incremental_builds(true);
            request["renderEnv"]["revisionPreview"] = serde_json::json!({revision_id: "accepted"});
            let output: serde_json::Value = serde_json::from_str(
                &engine
                    .layout_document_with_regions_json(&request.to_string())
                    .unwrap(),
            )
            .unwrap();
            assert!(!engine.pagination.borrow().last_incremental);
            let epoch = engine.display.borrow().binary_frame_epoch;
            let (_, converted) = with_layout_conversion(force_full, || {
                engine
                    .build_display_list_frame(&extras(&output), epoch)
                    .unwrap()
            });
            let mut snapshots = Vec::new();
            let record =
                || serde_json::to_value(engine.with_display_list(Clone::clone).unwrap()).unwrap();
            let list = engine.with_display_list(Clone::clone).unwrap();
            assert!(list.pages.len() >= 4);
            assert!(!list.pages[0].unbuilt && list.pages[1..].iter().all(|page| page.unbuilt));
            assert_eq!(converted, if force_full { list.pages.len() } else { 1 });
            assert!(list.pages[0].header.is_some() && list.pages[0].footer.is_some());
            assert!(!list.pages[0].note_areas.is_empty());
            {
                let pagination = engine.pagination.borrow();
                let pages = &pagination.layout.as_ref().unwrap().pages;
                assert!(
                    pages
                        .iter()
                        .filter(|page| page
                            .fragments
                            .iter()
                            .any(|f| matches!(f, Fragment::Table(_))))
                        .count()
                        > 1
                );
                assert!(pages.iter().any(|page| page.fragments.iter().any(
                    |f| matches!(f, Fragment::Shape(shape) if shape.is_anchored == Some(true))
                )));
                assert!(pages.iter().any(|page| page.size != pages[0].size));
                assert_eq!(pages[0].page_label.as_deref(), Some("i"));
                assert!(
                    pages.iter().any(|page| page.section_index == Some(1)
                        && page.page_label.as_deref() == Some("1"))
                );
            }
            snapshots.push(record());
            let complete = || {
                let rest = engine
                    .with_display_list(|list| {
                        list.pages
                            .iter()
                            .enumerate()
                            .filter_map(|(index, page)| page.unbuilt.then_some(index))
                            .collect::<Vec<_>>()
                    })
                    .unwrap();
                let epoch = engine.display.borrow().binary_frame_epoch;
                with_layout_conversion(force_full, || {
                    engine.build_display_pages_frame(&rest, epoch).unwrap()
                });
                assert_eq!(
                    record(),
                    serde_json::to_value(full_display_build(
                        &engine,
                        &engine.resident_region_display_extras().unwrap()
                    ))
                    .unwrap()
                );
            };
            complete();
            snapshots.push(record());
            insert_x(&engine, "body", 5);
            let before = engine.stats();
            let epoch = engine.display.borrow().binary_frame_epoch;
            let (_, converted) = with_layout_conversion(force_full, || {
                engine.apply_and_layout("body", epoch).unwrap()
            });
            let after = engine.stats();
            assert_eq!(
                after.incremental_display_builds,
                before.incremental_display_builds + 1
            );
            if !force_full {
                assert_eq!(
                    converted as u64,
                    after.rebuilt_display_pages - before.rebuilt_display_pages
                );
            }
            snapshots.push(record());
            complete();
            snapshots.push(record());
            let rest = engine
                .with_display_list(|list| (1..list.pages.len()).collect::<Vec<_>>())
                .unwrap();
            let epoch = engine.display.borrow().binary_frame_epoch;
            let (_, converted) = with_layout_conversion(force_full, || {
                engine.release_display_pages_frame(&rest, epoch).unwrap()
            });
            assert_eq!(converted, if force_full { rest.len() } else { 0 });
            snapshots.push(record());
            complete();
            snapshots.push(record());
            runs.push(snapshots);
            docx_layout::clear_measure_fonts();
        }
        assert_eq!(runs[0].len(), runs[1].len());
        for (step, (selective, legacy)) in runs[0].iter().zip(&runs[1]).enumerate() {
            assert_eq!(selective, legacy, "step {step}");
        }
    }

    #[test]
    fn unbuilt_display_pages_build_on_request_and_match_a_full_build() {
        let (engine, extras) = paged_filler_engine(205, 48);
        let full_build = |engine: &EngineSession| full_display_build(engine, &extras);

        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let lazy = engine.with_display_list(Clone::clone).unwrap();
        let last = lazy.pages.len() - 1;
        assert!(last >= 3, "the fixture must span several pages");
        assert!(!lazy.pages[0].unbuilt);
        let expected = full_build(&engine);
        for (page, full) in lazy.pages.iter().zip(&expected.pages).skip(1) {
            assert!(page.unbuilt && page.primitives.is_empty());
            assert_eq!((&page.width, &page.height), (&full.width, &full.height));
            assert_eq!(page.content_bounds, full.content_bounds);
        }

        // An edit on the first page shifts every later position; a page built
        // afterwards must show the shifted positions.
        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " typed",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let shifted = engine.with_display_list(Clone::clone).unwrap();
        for (page, before) in shifted.pages.iter().zip(&lazy.pages).skip(1) {
            if page.unbuilt {
                let [start, end] = before.position_span.unwrap();
                assert_eq!(page.position_span, Some([start + 6, end + 6]));
            }
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&[last, 1], epoch).unwrap();
        let built = engine.with_display_list(Clone::clone).unwrap();
        let expected = full_build(&engine);
        assert_eq!(built.pages[0], expected.pages[0]);
        assert_eq!(built.pages[1], expected.pages[1]);
        assert_eq!(built.pages[last], expected.pages[last]);
        assert!(built.pages[2].unbuilt);

        // A later full build evicts pages outside the window.
        engine
            .build_display_list_frame(&format!("{extras} "), epoch + 1)
            .unwrap();
        let rebuilt = engine.with_display_list(Clone::clone).unwrap();
        assert!(!rebuilt.pages[0].unbuilt);
        assert!(rebuilt.pages[last].unbuilt && rebuilt.pages[2].unbuilt);
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&[last], epoch).unwrap();
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages[last],
            full_build(&engine).pages[last]
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn windowed_placeholder_span_frames_match_with_local_lowering_on_and_off() {
        // Building an engine clears the thread's measurement fonts, so each runs its edits first.
        let runs: Vec<Vec<_>> = [false, true]
            .into_iter()
            .map(|enabled| {
                let (engine, extras) = paged_filler_engine(216, 48);
                engine.set_local_lowering(enabled);
                engine.set_display_window(Some(0..1));
                engine.set_windowed_incremental_builds(true);
                let full = engine.build_display_list_frame(&extras, 0).unwrap();
                let mut retained = std::collections::HashMap::new();
                let applied =
                    crate::frame_delta::apply_placeholder_test_frame(&full, &mut retained);
                assert!(applied.pages.len() >= 4);
                assert!(applied.pages[1..].iter().all(|page| page.unbuilt));
                ["x", "y", "z"]
                    .into_iter()
                    .enumerate()
                    .map(|(edit, text)| {
                        engine
                            .doc()
                            .insert_text(
                                &crate::EditCtx::local("", ""),
                                crate::Position::new("body", 3 + edit as u32),
                                text,
                                crate::FormatPolicy::Inherit,
                            )
                            .unwrap();
                        let before = engine.with_display_list(Clone::clone).unwrap();
                        let epoch = engine.display.borrow().binary_frame_epoch;
                        let bytes = engine.apply_and_layout("body", epoch).unwrap();
                        let next =
                            crate::frame_delta::apply_placeholder_test_frame(&bytes, &mut retained);
                        assert_eq!(next, engine.with_display_list(Clone::clone).unwrap());
                        let operation_count =
                            u32::from_le_bytes(bytes[52..56].try_into().unwrap()) as usize;
                        let ranges: Vec<_> = (0..operation_count)
                            .map(|op| {
                                crate::frame_delta::FRAME_HEADER_LEN
                                    + op * crate::frame_delta::PAGE_OP_LEN
                            })
                            .filter(|&record| {
                                bytes[record] == crate::frame_delta::PAGE_OP_SHIFT_RANGE
                            })
                            .collect();
                        assert_eq!(ranges.len(), 1);
                        let record = ranges[0];
                        for (index, page) in next.pages.iter().enumerate().skip(1) {
                            assert!(page.unbuilt);
                            let [start, end] = before.pages[index].position_span.unwrap();
                            assert_eq!(page.position_span, Some([start + 1, end + 1]));
                            let first = u32::from_le_bytes(
                                bytes[record + 4..record + 8].try_into().unwrap(),
                            ) as usize;
                            let count = u32::from_le_bytes(
                                bytes[record + 24..record + 28].try_into().unwrap(),
                            ) as usize;
                            assert!((first..first + count).contains(&index));
                        }
                        next
                    })
                    .collect()
            })
            .collect();
        assert_eq!(runs[0], runs[1]);
        docx_layout::clear_measure_fonts();
    }

    /// Resident frames decode to the display list through both host modes and lowering settings.
    #[test]
    fn resident_frames_decode_to_the_display_list_across_build_release_and_edits() {
        let mut runs = Vec::new();
        for enabled in [false, true] {
            let (engine, extras) = paged_filler_engine(217, 48);
            engine.set_local_lowering(enabled);
            let mut retained = std::collections::HashMap::new();
            let mut frames = Vec::new();
            let mut decode = |bytes: Vec<u8>| {
                let decoded =
                    crate::frame_delta::apply_placeholder_test_frame(&bytes, &mut retained);
                assert_eq!(
                    decoded,
                    engine.with_display_list(Clone::clone).unwrap(),
                    "local lowering {enabled}, frame {}",
                    frames.len()
                );
                frames.push(bytes);
                decoded
            };
            let epoch = || engine.display.borrow().binary_frame_epoch;
            let edit = |offset, text: &str| {
                let ctx = crate::EditCtx::local("", "");
                if text.is_empty() {
                    engine
                        .doc()
                        .delete_range(&ctx, crate::StoryRange::new("body", offset, offset + 1))
                        .unwrap();
                } else {
                    engine
                        .doc()
                        .insert_text(
                            &ctx,
                            crate::Position::new("body", offset),
                            text,
                            crate::FormatPolicy::Inherit,
                        )
                        .unwrap();
                }
                engine.apply_and_layout("body", epoch()).unwrap()
            };
            let paragraphs = engine.doc().paragraphs("body").unwrap();
            let tail = paragraphs[..paragraphs.len() - 1]
                .iter()
                .map(|paragraph| u32::try_from(paragraph.text.encode_utf16().count()).unwrap() + 1)
                .sum::<u32>()
                + 3;

            engine.set_display_window(Some(0..1));
            engine.set_windowed_incremental_builds(true);
            let initial = decode(engine.build_display_list_frame(&extras, 0).unwrap());
            assert!(initial.pages.len() >= 4);
            assert!(!initial.pages[0].unbuilt);
            assert!(initial.pages[1..].iter().all(|page| page.unbuilt));
            let last = initial.pages.len() - 1;
            let distant: Vec<_> = (1..initial.pages.len()).collect();
            for (batch, pages) in distant.chunks(2).enumerate() {
                engine.set_windowed_incremental_builds(false);
                let built = decode(engine.build_display_pages_frame(pages, epoch()).unwrap());
                assert!(pages.iter().all(|&index| !built.pages[index].unbuilt));
                if batch == 0 {
                    assert!(built.pages[last].unbuilt);
                    decode(edit(tail, "x"));
                    decode(edit(tail, ""));
                }
            }
            assert!(
                engine
                    .with_display_list(|list| list.pages.iter().all(|page| !page.unbuilt))
                    .unwrap()
            );
            decode(edit(tail, "y"));

            engine.set_windowed_incremental_builds(true);
            let mut before = decode(
                engine
                    .release_display_pages_frame(&distant, epoch())
                    .unwrap(),
            );
            assert!(before.pages[1..].iter().all(|page| page.unbuilt));
            for (offset, text) in ["a", "b", "c"].into_iter().enumerate() {
                let bytes = edit(3 + offset as u32, text);
                let operation_count =
                    u32::from_le_bytes(bytes[52..56].try_into().unwrap()) as usize;
                assert_eq!(
                    (0..operation_count)
                        .filter(|&op| {
                            bytes[crate::frame_delta::FRAME_HEADER_LEN
                                + op * crate::frame_delta::PAGE_OP_LEN]
                                == crate::frame_delta::PAGE_OP_SHIFT_RANGE
                        })
                        .count(),
                    1
                );
                let next = decode(bytes);
                for (page, previous) in next.pages.iter().zip(&before.pages).skip(1) {
                    assert!(page.unbuilt);
                    let [start, end] = previous.position_span.unwrap();
                    assert_eq!(page.position_span, Some([start + 1, end + 1]));
                }
                before = next;
            }
            engine.set_windowed_incremental_builds(false);
            let rebuilt = decode(engine.build_display_pages_frame(&[last], epoch()).unwrap());
            assert!(!rebuilt.pages[last].unbuilt);
            assert!(rebuilt.pages[1..last].iter().all(|page| page.unbuilt));
            engine.set_display_window(Some(last..last + 1));
            engine.set_windowed_incremental_builds(true);
            let deleted = decode(edit(tail + 3, ""));
            assert!(!deleted.pages[last].unbuilt);

            engine.set_windowed_incremental_builds(false);
            decode(edit(tail + 3, "z"));
            runs.push(frames);
            docx_layout::clear_measure_fonts();
        }
        assert_eq!(runs[0].len(), runs[1].len());
        for (step, (disabled, enabled)) in runs[0].iter().zip(&runs[1]).enumerate() {
            assert_eq!(disabled, enabled, "frame {step}");
        }
    }

    #[test]
    fn released_display_pages_keep_ids_and_rebuild_after_an_incremental_edit() {
        let (engine, extras) = paged_filler_engine(214, 48);
        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let last = engine
            .with_display_list(|list| list.pages.len() - 1)
            .unwrap();
        assert!(last >= 3);
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&[1, last], epoch).unwrap();
        let before = engine.stats();
        let old_snapshots = engine.display.borrow().pages.clone();
        let expected_placeholders = {
            let pagination = engine.pagination.borrow();
            docx_layout::build_resident_display_list_partial_observed(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                &extras,
                &|_| false,
                &mut || {},
            )
            .unwrap()
            .1
        };
        let epoch = engine.display.borrow().binary_frame_epoch;
        let delta = engine
            .release_display_pages_frame(&[last, 1, 1, 2], epoch)
            .unwrap();
        let after = engine.stats();
        assert_eq!(after.doc_epoch, before.doc_epoch);
        assert_eq!(after.layout_epoch, before.layout_epoch);
        assert_eq!(after.frame_epoch, before.frame_epoch + 1);
        assert_eq!(after.display_builds, before.display_builds);
        assert_eq!(u32::from_le_bytes(delta[12..16].try_into().unwrap()), 0);
        assert_eq!(u32::from_le_bytes(delta[52..56].try_into().unwrap()), 2);
        for (operation, index) in [1, last].into_iter().enumerate() {
            let start =
                crate::frame_delta::FRAME_HEADER_LEN + operation * crate::frame_delta::PAGE_OP_LEN;
            assert_eq!(delta[start], crate::frame_delta::PAGE_OP_UPSERT);
            assert_eq!(
                u32::from_le_bytes(delta[start + 4..start + 8].try_into().unwrap()),
                index as u32
            );
            assert_eq!(
                u64::from_le_bytes(delta[start + 8..start + 16].try_into().unwrap()),
                old_snapshots[index].page_id
            );
            assert_eq!(
                u32::from_le_bytes(delta[start + 24..start + 28].try_into().unwrap()),
                0
            );
        }
        let released = engine.with_display_list(Clone::clone).unwrap();
        for index in [1, last] {
            assert_eq!(released.pages[index], expected_placeholders.pages[index]);
            assert!(released.pages[index].position_span.is_some());
        }
        {
            let display = engine.display.borrow();
            for (index, snapshot) in display.pages.iter().enumerate() {
                assert_eq!(snapshot.page_id, old_snapshots[index].page_id);
                if [1, last].contains(&index) {
                    assert!(snapshot.primitive_ids.is_empty());
                    assert!(snapshot.positions.is_empty());
                    assert!(snapshot.note_anchors.is_empty());
                }
            }
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&[last], epoch).unwrap();
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages[last],
            full_display_build(&engine, &extras).pages[last]
        );

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " typed",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let edited_delta = engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(&edited_delta[0..4], b"FDV1");
        assert_eq!(
            u64::from_le_bytes(edited_delta[40..48].try_into().unwrap()),
            epoch
        );
        let expected = full_display_build(&engine, &extras);
        let edited = engine.with_display_list(Clone::clone).unwrap();
        for (page, full) in edited.pages.iter().zip(&expected.pages) {
            if !page.unbuilt {
                assert_eq!(page, full);
            }
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        let all: Vec<usize> = (0..edited.pages.len()).collect();
        engine.build_display_pages_frame(&all, epoch).unwrap();
        assert_eq!(engine.with_display_list(Clone::clone).unwrap(), expected);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn superseded_display_page_releases_do_not_mutate_retained_state() {
        let (engine, extras) = paged_filler_engine(215, 40);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let before = engine.with_display_list(Clone::clone).unwrap();
        let snapshots = engine.display.borrow().pages.clone();
        let next_page_id = engine.display.borrow().next_page_id;
        let resident = format!("{:?}", engine.display.borrow().resident_input);
        let assert_unchanged = || {
            assert_eq!(engine.with_display_list(Clone::clone).unwrap(), before);
            let display = engine.display.borrow();
            assert_eq!(display.pages, snapshots);
            assert_eq!(display.frame_epoch, epoch);
            assert_eq!(display.binary_frame_epoch, epoch);
            assert_eq!(display.next_page_id, next_page_id);
            assert_eq!(format!("{:?}", display.resident_input), resident);
        };
        for stale in [epoch - 1, epoch + 1] {
            assert!(
                engine
                    .release_display_pages_frame(&[1], stale)
                    .unwrap()
                    .is_empty()
            );
            assert_unchanged();
        }
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 0),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        assert_ne!(
            engine.doc_epoch(),
            engine.display.borrow().encoded_doc_epoch
        );
        assert!(
            engine
                .release_display_pages_frame(&[1], epoch)
                .unwrap()
                .is_empty()
        );
        assert_unchanged();
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn a_layout_replacement_supersedes_display_page_release() {
        let (engine, extras) = paged_filler_engine(216, 40);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let before = engine.with_display_list(Clone::clone).unwrap();
        let snapshots = engine.display.borrow().pages.clone();
        let input =
            serde_json::to_string(engine.pagination.borrow().input.as_ref().unwrap()).unwrap();
        engine.layout_document_json(&input).unwrap();
        assert_ne!(
            engine.stats().layout_epoch,
            engine.display.borrow().encoded_layout_epoch
        );
        assert!(
            engine
                .release_display_pages_frame(&[1], epoch)
                .unwrap()
                .is_empty()
        );
        assert_eq!(engine.with_display_list(Clone::clone).unwrap(), before);
        assert_eq!(engine.display.borrow().pages, snapshots);
        assert_eq!(engine.display.borrow().binary_frame_epoch, epoch);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_that_moves_every_later_page_leaves_unbuilt_pages_unbuilt() {
        let (engine, extras) = paged_filler_engine(206, 40);
        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert!(pages >= 4, "the fixture must span several pages");

        // A second line in the first paragraph moves the start of every later page.
        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " that now runs long enough to wrap onto a second line of the page",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let incremental_builds = engine.stats().incremental_display_builds;
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            incremental_builds + 1
        );
        assert_eq!(engine.pagination.borrow().rebuilt_page_end, pages);

        let edited = engine.with_display_list(Clone::clone).unwrap();
        let windowed = {
            let pagination = engine.pagination.borrow();
            docx_layout::build_resident_display_list_partial_observed(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                &extras,
                &|index| index == 0,
                &mut || {},
            )
            .unwrap()
            .1
        };
        assert_eq!(edited.pages, windowed.pages);
        assert!(edited.pages[1..].iter().all(|page| page.unbuilt));

        let epoch = engine.display.borrow().binary_frame_epoch;
        let rest: Vec<usize> = (1..pages).collect();
        engine.build_display_pages_frame(&rest, epoch).unwrap();
        let built = engine.with_display_list(Clone::clone).unwrap();
        assert_eq!(built.pages, full_display_build(&engine, &extras).pages);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_turns_re_placed_pages_outside_the_window_into_placeholders() {
        use docx_layout::display_list::resident_conversion_test_support::with_layout_conversion;

        let (engine, extras) = paged_filler_engine(207, 40);
        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert!(pages >= 4, "the fixture must span several pages");
        let rest: Vec<usize> = (1..pages).collect();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&rest, epoch).unwrap();
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages,
            full_display_build(&engine, &extras).pages
        );

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " that now runs long enough to wrap onto a second line of the page",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds + 1
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            1
        );
        assert_eq!(engine.pagination.borrow().rebuilt_page_start, 0);
        assert_eq!(engine.pagination.borrow().rebuilt_page_end, pages);

        let edited = engine.with_display_list(Clone::clone).unwrap();
        let windowed = {
            let pagination = engine.pagination.borrow();
            with_layout_conversion(true, || {
                docx_layout::build_resident_display_list_partial_observed(
                    pagination.input.as_ref().unwrap(),
                    pagination.layout.as_ref().unwrap(),
                    &extras,
                    &|index| index == 0,
                    &mut || {},
                )
                .unwrap()
                .1
            })
            .0
        };
        assert_eq!(
            serde_json::to_value(&edited).unwrap(),
            serde_json::to_value(&windowed).unwrap()
        );
        assert!(!edited.pages[0].unbuilt);
        assert_eq!(
            edited.pages[0],
            full_display_build(&engine, &extras).pages[0]
        );
        for (page, placeholder) in edited.pages.iter().zip(&windowed.pages).skip(1) {
            assert!(page.unbuilt && page.primitives.is_empty());
            assert!(page.position_span.is_some());
            assert_eq!(page.position_span, placeholder.position_span);
            assert_eq!(page, placeholder);
        }

        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&rest, epoch).unwrap();
        let built = engine.with_display_list(Clone::clone).unwrap();
        assert_eq!(built.pages, full_display_build(&engine, &extras).pages);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_builds_the_caret_page_outside_the_display_window() {
        use yrs::{Assoc, IndexedSequence};

        let (engine, extras) = paged_filler_engine(209, 160);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let initial = engine.with_display_list(Clone::clone).unwrap();
        assert!(initial.pages.len() >= 11);
        assert!(initial.pages.iter().all(|page| !page.unbuilt));

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, "body").unwrap();
        let head = text.sticky_index(&txn, offset, Assoc::After).unwrap();
        drop(txn);
        engine.set_resident_caret_head(Some(("body".to_owned(), head)));
        assert_eq!(
            engine
                .resident_caret_snapshot(Some((&paragraph.para_id, offset)))
                .unwrap()
                .caret_rect
                .unwrap()
                .page_index,
            0
        );
        engine.set_display_window(Some(8..11));
        engine.set_windowed_incremental_builds(true);

        let insertion = "x";
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                insertion,
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds + 1
        );
        let edited = engine.with_display_list(Clone::clone).unwrap();
        assert!(!edited.pages[0].unbuilt);
        assert!(edited.pages[8..11].iter().all(|page| !page.unbuilt));
        let caret = engine
            .resident_caret_snapshot(Some((
                &paragraph.para_id,
                offset + u32::try_from(insertion.encode_utf16().count()).unwrap(),
            )))
            .unwrap();
        assert_eq!(
            caret.frame_epoch,
            engine.display.borrow().binary_frame_epoch
        );
        assert_eq!(caret.caret_rect.unwrap().page_index, 0);
        assert_eq!(
            edited.pages[0],
            full_display_build(&engine, &extras).pages[0]
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn a_full_windowed_build_keeps_only_the_window_and_the_caret_page() {
        use docx_layout::display_list::resident_conversion_test_support::with_layout_conversion;
        use yrs::{Assoc, IndexedSequence};

        let (engine, extras) = paged_filler_engine(213, 160);
        engine.build_display_list_frame(&extras, 0).unwrap();
        assert!(!engine.pagination.borrow().last_incremental);
        let full = full_display_build(&engine, &extras);
        assert!(full.pages.len() >= 11);

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, "body").unwrap();
        let head = text.sticky_index(&txn, 0, Assoc::After).unwrap();
        drop(txn);
        engine.set_resident_caret_head(Some(("body".to_owned(), head)));
        engine.set_display_window(Some(8..11));
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_list_frame(&extras, epoch).unwrap();
        assert!(
            engine
                .with_display_list(|list| list.pages.iter().all(|page| !page.unbuilt))
                .unwrap(),
            "without windowed builds a full build keeps every built page"
        );

        engine.set_windowed_incremental_builds(true);
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_list_frame(&extras, epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds
        );
        let windowed = engine.with_display_list(Clone::clone).unwrap();
        for (index, page) in windowed.pages.iter().enumerate() {
            assert_eq!(page.unbuilt, !(8..11).contains(&index), "page {index}");
        }

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 0),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let (_, converted) = with_layout_conversion(false, || {
            engine
                .build_display_list_frame(&format!("{extras} "), epoch)
                .unwrap()
        });
        assert_eq!(converted, 4);
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds
        );
        let full = full_display_build(&engine, &extras);
        let windowed = engine.with_display_list(Clone::clone).unwrap();
        let legacy = {
            let pagination = engine.pagination.borrow();
            with_layout_conversion(true, || {
                docx_layout::build_resident_display_list_partial_observed(
                    pagination.input.as_ref().unwrap(),
                    pagination.layout.as_ref().unwrap(),
                    &extras,
                    &|index| index == 0 || (8..11).contains(&index),
                    &mut || {},
                )
                .unwrap()
                .1
            })
            .0
        };
        assert_eq!(
            serde_json::to_value(&windowed).unwrap(),
            serde_json::to_value(&legacy).unwrap()
        );
        for (index, (page, full_page)) in windowed.pages.iter().zip(&full.pages).enumerate() {
            if index == 0 || (8..11).contains(&index) {
                assert_eq!(page, full_page, "page {index} is built");
            } else {
                assert!(
                    page.unbuilt && page.primitives.is_empty(),
                    "page {index} waits"
                );
                assert_eq!(page.width, full_page.width);
                assert_eq!(page.height, full_page.height);
                assert_eq!(page.content_bounds, full_page.content_bounds);
            }
        }
        assert_eq!(
            engine
                .resident_caret_snapshot(Some((&paragraph.para_id, 1)))
                .unwrap()
                .caret_rect
                .unwrap()
                .page_index,
            0
        );

        let rest: Vec<usize> = (0..full.pages.len())
            .filter(|index| *index != 0 && !(8..11).contains(index))
            .collect();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_pages_frame(&rest, epoch).unwrap();
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages,
            full.pages
        );
        engine.set_display_retain_built_pages(true);
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine
            .build_display_list_frame(&format!("{extras}  "), epoch)
            .unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds
        );
        let retained = engine.with_display_list(Clone::clone).unwrap();
        assert!(retained.pages.iter().all(|page| !page.unbuilt));
        assert_eq!(retained.pages, full_display_build(&engine, &extras).pages);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_builds_the_table_cell_caret_page_outside_the_display_window() {
        use yrs::{Assoc, IndexedSequence};

        let (engine, extras, _) = paged_region_engine(
            211,
            &format!(
                r#"<w:tbl><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Editable cell paragraph</w:t></w:r></w:p></w:tc></w:tr></w:tbl>{}"#,
                "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(156)
            ),
        );
        let ctx = crate::EditCtx::local("", "");
        let cell_story = "body:t0:r0c0".to_owned();
        engine.build_display_list_frame(&extras, 0).unwrap();
        let initial = engine.with_display_list(Clone::clone).unwrap();
        assert!(initial.pages.len() >= 11);
        assert!(initial.pages.iter().all(|page| !page.unbuilt));
        assert!(matches!(
            engine.pagination.borrow().layout.as_ref().unwrap().pages[0]
                .fragments
                .first(),
            Some(Fragment::Table(_))
        ));

        let paragraph = engine.doc().paragraphs(&cell_story).unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, &cell_story).unwrap();
        let head = text.sticky_index(&txn, offset, Assoc::After).unwrap();
        drop(txn);
        engine.set_resident_caret_head(Some((cell_story.clone(), head)));
        engine.set_display_window(Some(8..11));
        engine.set_windowed_incremental_builds(true);

        engine
            .doc()
            .insert_text(
                &ctx,
                crate::Position::new(&cell_story, offset),
                " that now runs long enough to wrap onto a second line of the page",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout(&cell_story, epoch).unwrap();
        // A page-count change takes the full region pass by design.
        assert_eq!(
            engine
                .pagination
                .borrow()
                .layout
                .as_ref()
                .unwrap()
                .pages
                .len(),
            initial.pages.len()
        );
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds + 1
        );
        let edited = engine.with_display_list(Clone::clone).unwrap();
        assert_eq!(engine.pagination.borrow().rebuilt_page_start, 0);
        assert_eq!(
            engine.pagination.borrow().rebuilt_page_end,
            edited.pages.len()
        );
        assert!(!edited.pages[0].unbuilt);
        assert!(edited.pages[8..11].iter().all(|page| !page.unbuilt));
        assert!(edited.pages[1..8].iter().all(|page| page.unbuilt));
        assert_eq!(
            edited.pages[0],
            full_display_build(&engine, &extras).pages[0]
        );

        // A picture in the caret's paragraph keeps it mapped.
        let mut input = engine.pagination.borrow().input.clone().unwrap();
        let LayoutBlock::Table(table) = &mut input.measured[0].block else {
            panic!("expected the table");
        };
        let LayoutBlock::Paragraph(cell) = &mut table.rows[0].cells[0].blocks[0] else {
            panic!("expected the cell paragraph");
        };
        let span = (cell.pm_start.unwrap(), cell.pm_end.unwrap());
        cell.runs.push(Run::Image(
            serde_json::from_value(serde_json::json!({"src": "picture", "width": 8, "height": 8}))
                .unwrap(),
        ));
        assert_eq!(
            resident_paragraph_span(&input, &paragraph.para_id),
            Some(span)
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn a_caret_maps_only_where_every_span_agrees() {
        let span = |pm_start, pm_end, raw_start, raw_end, atom| crate::bridge::SourceSpan {
            pm_start,
            pm_end,
            paragraph: 0,
            raw_start,
            raw_end,
            atom,
        };
        let text = [span(1, 5, 0, 4, false), span(5, 8, 4, 7, false)];
        assert_eq!(span_position(text.iter(), 2), Some(3));
        assert_eq!(span_position(text.iter(), 4), Some(5));
        assert_eq!(span_position(text.iter(), 7), Some(8));
        let chart = [span(1, 2, 0, 1, true), span(2, 6, 1, 5, false)];
        assert_eq!(span_position(chart.iter(), 0), Some(1));
        assert_eq!(span_position(chart.iter(), 1), Some(2));
        let control = [
            span(5, 7, 4, 5, false),
            span(7, 9, 4, 5, false),
            span(10, 12, 5, 7, false),
        ];
        assert_eq!(span_position(control.iter(), 5), None);
        let wide = [span(1, 2, 0, 3, true)];
        assert_eq!(span_position(wide.iter(), 1), None);
        assert_eq!(span_position(wide.iter(), 3), Some(2));
    }

    #[test]
    fn windowed_builds_keep_the_caret_page_or_every_re_placed_page() {
        let (engine, _) = paged_filler_engine(213, 160);
        let pagination = engine.pagination.borrow();
        let layout = pagination.layout.as_ref().unwrap();
        let display = DisplayState {
            window: Some(8..11),
            windowed_incremental_builds: true,
            ..DisplayState::default()
        };
        let rebuilt: HashSet<usize> = (0..layout.pages.len()).collect();
        let built = |caret| {
            let build = window_build_pages(&display, layout, rebuilt.iter().copied(), caret);
            (0..layout.pages.len())
                .filter(|index| build.as_ref().is_none_or(|pages| pages.contains(index)))
                .collect::<Vec<_>>()
        };
        let Some(Fragment::Paragraph(last)) = layout.pages[0].fragments.last() else {
            panic!("expected a paragraph at the end of the first page");
        };
        let start = last.pm_start.unwrap();
        assert_eq!(built(None), [8, 9, 10]);
        assert_eq!(
            built(Some(CaretExtent::Position(start + 1.0))),
            [0, 8, 9, 10]
        );
        assert_eq!(
            built(Some(CaretExtent::Unmapped)),
            (0..layout.pages.len()).collect::<Vec<_>>()
        );
        drop(pagination);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_builds_the_caret_page_after_a_long_inline_content_control() {
        use yrs::{Assoc, IndexedSequence};

        let (engine, extras) = paged_engine(
            212,
            &format!(
                "<w:p><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>{}</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>Trailing text</w:t></w:r></w:p>{}",
                "Content control text that runs over several pages. ".repeat(40),
                "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(160)
            ),
        );
        engine.build_display_list_frame(&extras, 0).unwrap();
        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, "body").unwrap();
        let bounds = crate::op::para_bounds(&text, &txn).remove(0);
        let head = text
            .sticky_index(&txn, bounds.pilcrow, Assoc::After)
            .unwrap();
        drop(txn);
        let (_, paragraph_end) = resident_paragraph_span(
            engine.pagination.borrow().input.as_ref().unwrap(),
            &bounds.para_id,
        )
        .unwrap();
        let last_page = |engine: &EngineSession| {
            let pagination = engine.pagination.borrow();
            let pages = &pagination.layout.as_ref().unwrap().pages;
            pages
                .iter()
                .rposition(|page| {
                    page.fragments.iter().any(|fragment| {
                        matches!(fragment, Fragment::Paragraph(value)
                            if value.pm_start.is_some_and(|start| start <= paragraph_end))
                    })
                })
                .unwrap()
        };
        let caret_page = last_page(&engine);
        assert!(caret_page >= 2, "the paragraph must run over several pages");
        engine.set_resident_caret_head(Some(("body".to_owned(), head)));
        engine.set_display_window(Some(caret_page + 4..caret_page + 7));
        engine.set_windowed_incremental_builds(true);

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", bounds.pilcrow),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(last_page(&engine), caret_page);
        let edited = engine.with_display_list(Clone::clone).unwrap();
        assert!(!edited.pages[caret_page].unbuilt);
        assert_eq!(
            edited.pages[caret_page],
            full_display_build(&engine, &extras).pages[caret_page]
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_builds_the_caret_page_when_equal_paragraph_lengths_hide_a_shift() {
        use yrs::{Assoc, IndexedSequence};

        let (engine, extras, _) = paged_region_engine(
            214,
            &format!(
                "<w:p>{}<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>Z</w:t></w:r></w:sdtContent></w:sdt></w:p>{}",
                "<w:r><w:t>Text across several pages</w:t><w:br/></w:r>".repeat(80),
                "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(40)
            ),
        );
        engine.build_display_list_frame(&extras, 0).unwrap();
        let ctx = crate::EditCtx::local("", "");
        for index in 0..2 {
            engine
                .doc()
                .insert_embed(
                    &ctx,
                    crate::Position::new("body", index),
                    "chart",
                    vec![
                        ("chartJson".to_owned(), Any::from("{}")),
                        ("width".to_owned(), Any::Number(8.0)),
                        ("height".to_owned(), Any::Number(8.0)),
                    ],
                )
                .unwrap();
        }
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let initial = engine.with_display_list(Clone::clone).unwrap();
        assert!(initial.pages.len() >= 4);
        assert!(initial.pages.iter().all(|page| !page.unbuilt));
        assert_eq!(initial.pages, full_display_build(&engine, &extras).pages);

        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, "body").unwrap();
        let bounds = crate::op::para_bounds(&text, &txn).remove(0);
        drop(txn);
        let (start, end) = resident_paragraph_span(
            engine.pagination.borrow().input.as_ref().unwrap(),
            &bounds.para_id,
        )
        .unwrap();
        assert_eq!(start, 2.0);
        assert_eq!(end - start, f64::from(bounds.len()) + 2.0);

        let caret_page = 1;
        let fragment_end = |engine: &EngineSession, page: usize| {
            engine.pagination.borrow().layout.as_ref().unwrap().pages[page]
                .fragments
                .iter()
                .find_map(|fragment| match fragment {
                    Fragment::Paragraph(value) if block_key(&value.block_id) == bounds.para_id => {
                        value.pm_end
                    }
                    _ => None,
                })
                .unwrap()
        };
        assert!(fragment_end(&engine, caret_page + 1) < end);
        let offset = (fragment_end(&engine, caret_page) - 2.0) as u32;
        assert!(offset > 2 && offset < bounds.pilcrow - 1);
        let txn = engine.doc().yrs_doc().transact();
        let text = crate::story_ref(&txn, "body").unwrap();
        let head = text.sticky_index(&txn, offset, Assoc::After).unwrap();
        drop(txn);
        engine.set_resident_caret_head(Some(("body".to_owned(), head)));
        let window_page = initial.pages.len() - 1;
        engine.set_display_window(Some(window_page..window_page + 1));
        engine.set_windowed_incremental_builds(true);

        engine
            .doc()
            .insert_text(
                &ctx,
                crate::Position::new("body", offset),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds + 1
        );
        let txn = engine.doc().yrs_doc().transact();
        let caret_index = {
            let caret = engine.resident_caret_head.borrow();
            caret.as_ref().unwrap().1.get_offset(&txn).unwrap().index
        };
        drop(txn);
        assert_eq!(caret_index, offset + 1);
        let caret_position = f64::from(caret_index) + 1.0;
        assert_eq!(caret_position, fragment_end(&engine, caret_page) - 1.0);
        let (start, end) = resident_paragraph_span(
            engine.pagination.borrow().input.as_ref().unwrap(),
            &bounds.para_id,
        )
        .unwrap();
        assert_eq!(end - start, f64::from(bounds.len() + 1) + 2.0);
        assert_eq!(start + 1.0 + f64::from(caret_index), caret_position + 2.0);
        let edited = engine.with_display_list(Clone::clone).unwrap();
        assert_eq!(edited.pages.len(), initial.pages.len());
        assert!(!edited.pages[caret_page].unbuilt);
        assert!(!edited.pages[window_page].unbuilt);
        assert!(edited.pages[caret_page + 1].unbuilt);
        assert_eq!(
            edited.pages[caret_page],
            full_display_build(&engine, &extras).pages[caret_page]
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_with_windowed_incremental_builds_disabled_builds_every_re_placed_page() {
        for opt_in in [false, true] {
            let (engine, extras) = paged_filler_engine(210, 40);
            assert!(!engine.display.borrow().windowed_incremental_builds);
            engine.set_display_window(Some(0..1));
            engine.build_display_list_frame(&extras, 0).unwrap();
            let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
            assert!(pages >= 4);
            let rest: Vec<usize> = (1..pages).collect();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.build_display_pages_frame(&rest, epoch).unwrap();
            if opt_in {
                engine.set_windowed_incremental_builds(true);
                engine.set_windowed_incremental_builds(false);
            }

            let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
            let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", offset),
                    " that now runs long enough to wrap onto a second line of the page",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
            let before = engine.stats();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.apply_and_layout("body", epoch).unwrap();
            let after = engine.stats();
            assert_eq!(
                after.incremental_display_builds,
                before.incremental_display_builds + 1
            );
            assert_eq!(
                after.rebuilt_display_pages - before.rebuilt_display_pages,
                pages as u64
            );
            assert_eq!(engine.pagination.borrow().rebuilt_page_start, 0);
            assert_eq!(engine.pagination.borrow().rebuilt_page_end, pages);
            let edited = engine.with_display_list(Clone::clone).unwrap();
            assert!(edited.pages.iter().all(|page| !page.unbuilt));
            assert_eq!(edited.pages, full_display_build(&engine, &extras).pages);
            docx_layout::clear_measure_fonts();
        }
    }

    #[test]
    fn an_edit_without_a_display_window_builds_every_re_placed_page() {
        let (engine, extras) = paged_filler_engine(208, 40);
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert!(pages >= 4, "the fixture must span several pages");

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " that now runs long enough to wrap onto a second line of the page",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let before = engine.stats();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let after = engine.stats();
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds + 1
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            pages as u64
        );
        assert_eq!(engine.pagination.borrow().rebuilt_page_start, 0);
        assert_eq!(engine.pagination.borrow().rebuilt_page_end, pages);
        let edited = engine.with_display_list(Clone::clone).unwrap();
        assert!(edited.pages.iter().all(|page| !page.unbuilt));
        assert_eq!(edited.pages, full_display_build(&engine, &extras).pages);
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn edits_pages_apart_rebuild_only_their_own_pages() {
        let (engine, extras) = paged_filler_engine(207, 48);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert!(pages >= 4, "the fixture must span several pages");

        // A word at the end of the second and the last paragraph wraps neither.
        let ends: Vec<u32> = engine
            .doc()
            .paragraphs("body")
            .unwrap()
            .iter()
            .scan(0_u32, |start, paragraph| {
                let length = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
                let end = *start + length;
                *start = end + 1;
                Some(end)
            })
            .collect();
        for end in [ends[ends.len() - 1], ends[1]] {
            engine
                .doc()
                .insert_text(
                    &crate::EditCtx::local("", ""),
                    crate::Position::new("body", end),
                    " too",
                    crate::FormatPolicy::Inherit,
                )
                .unwrap();
        }
        let incremental_builds = engine.stats().incremental_display_builds;
        let rebuilt_display_pages = engine.stats().rebuilt_display_pages;
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            incremental_builds + 1
        );
        assert_eq!(
            engine.pagination.borrow().rebuilt_page_ranges,
            vec![0..1, pages - 1..pages]
        );
        assert_eq!(engine.stats().rebuilt_pages, 2);
        assert_eq!(
            engine.stats().rebuilt_display_pages,
            rebuilt_display_pages + 2
        );
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages,
            full_display_build(&engine, &extras).pages
        );
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn a_retained_page_whose_stamps_changed_is_rebuilt() {
        let (engine, extras, _) = paged_region_engine(
            208,
            &format!(
                "<w:p><w:r><w:t>Editable paragraph</w:t></w:r></w:p>{}",
                "<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(48)
            ),
        );
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert!(pages >= 4, "the fixture must span several pages");
        let last = pages - 1;
        engine
            .pagination
            .borrow_mut()
            .layout
            .as_mut()
            .unwrap()
            .pages[last]
            .page_label = Some("i".to_owned());
        engine.display.borrow_mut().list.as_mut().unwrap().pages[last].page_label =
            Some("i".to_owned());

        let paragraph = engine.doc().paragraphs("body").unwrap().remove(0);
        let offset = u32::try_from(paragraph.text.encode_utf16().count()).unwrap();
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", offset),
                " too",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let incremental_builds = engine.stats().incremental_display_builds;
        let epoch = engine.display.borrow().binary_frame_epoch;
        assert!(
            engine
                .apply_and_layout_regions_resident("body", &mut |_| {})
                .unwrap()
        );
        assert!(
            engine
                .pagination
                .borrow()
                .restamped_pages
                .as_ref()
                .unwrap()
                .contains(&last)
        );
        engine.build_display_list_frame(&extras, epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            incremental_builds + 1
        );
        assert!(
            engine
                .pagination
                .borrow()
                .rebuilt_page_ranges
                .iter()
                .all(|range| !range.contains(&last))
        );
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap().pages,
            full_display_build(&engine, &extras).pages
        );
        docx_layout::clear_measure_fonts();
    }

    /// Unchanged page stamps leave an empty tracked set after a frame.
    #[test]
    fn an_edit_without_stamp_changes_clears_restamped_pages_after_the_frame() {
        let (engine, extras, _) = paged_region_engine(
            220,
            &"<w:p><w:r><w:t>Filler paragraph</w:t></w:r></w:p>".repeat(48),
        );
        engine.build_display_list_frame(&extras, 0).unwrap();
        let initial = engine.with_display_list(Clone::clone).unwrap();
        engine
            .edit_resident_text(crate::StoryRange::new("body", 3, 3), Some("x"), true)
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        let pagination = engine.pagination.borrow();
        assert!(pagination.last_incremental);
        assert_eq!(pagination.restamped_pages, Some(BTreeSet::new()));
        let layout = pagination.layout.as_ref().unwrap();
        assert_eq!(initial.pages.len(), layout.pages.len());
        assert!(
            initial
                .pages
                .iter()
                .zip(&layout.pages)
                .all(|(shown, page)| page_stamps_match(shown, page))
        );
        drop(pagination);
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap(),
            full_display_build(&engine, &extras),
        );
        docx_layout::clear_measure_fonts();
    }

    /// Page count changes invalidate tracking until the full display build.
    #[test]
    fn a_page_count_change_makes_restamped_pages_unknown_before_the_frame() {
        let (engine, extras) = paged_filler_engine(221, 48);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
        assert_eq!(
            engine.pagination.borrow().restamped_pages,
            Some(BTreeSet::new()),
        );
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", 3, 3),
                Some(&" additional text".repeat(240)),
                true,
            )
            .unwrap();
        let resident = engine.resident_layout_input("body").unwrap();
        engine
            .layout_document_value_with_fingerprints(
                resident.input,
                resident.block_fingerprints,
                None,
                false,
            )
            .unwrap();
        let pagination = engine.pagination.borrow();
        assert!(pagination.last_incremental);
        assert!(pagination.layout.as_ref().unwrap().pages.len() > pages);
        assert_eq!(pagination.restamped_pages, None);
        drop(pagination);
        let incremental_builds = engine.stats().incremental_display_builds;
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.build_display_list_frame(&extras, epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            incremental_builds
        );
        assert_eq!(
            engine.pagination.borrow().restamped_pages,
            Some(BTreeSet::new()),
        );
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap(),
            full_display_build(&engine, &extras),
        );
        docx_layout::clear_measure_fonts();
    }

    /// An engine laid out over small pages whose one table row splits across several.
    fn split_row_engine(client_id: u64) -> (EngineSession, String) {
        docx_layout::clear_measure_fonts();
        let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
        let engine = EngineSession::new(client_id);
        let cell: String = (0..60)
            .map(|index| format!("<w:p><w:r><w:t>Tall cell line {index}</w:t></w:r></w:p>"))
            .collect();
        let body = format!(
            "<w:p><w:r><w:t>Before the table</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w=\"3600\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3600\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"3600\" w:type=\"dxa\"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl><w:p><w:r><w:t>After the table</w:t></w:r></w:p>"
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
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let extras =
            serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
        (engine, extras)
    }

    #[test]
    fn an_unbuilt_page_spans_every_position_a_split_row_places_on_it() {
        let (engine, extras) = split_row_engine(206);
        engine.set_display_window(Some(0..1));
        engine.build_display_list_frame(&extras, 0).unwrap();
        let lazy = engine.with_display_list(Clone::clone).unwrap();
        let full = {
            let pagination = engine.pagination.borrow();
            docx_layout::build_display_list_value_from_resident(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                &extras,
            )
            .unwrap()
        };
        let unbuilt: Vec<usize> = (0..lazy.pages.len())
            .filter(|&index| lazy.pages[index].unbuilt)
            .collect();
        assert!(
            unbuilt.len() >= 4,
            "the row must split across several unbuilt pages"
        );
        for &index in &unbuilt {
            let page = serde_json::to_value(&full.pages[index]).unwrap();
            for primitive in page["primitives"].as_array().unwrap() {
                let Some(position) = primitive["docStart"].as_i64() else {
                    continue;
                };
                assert!(
                    lazy.pages[index]
                        .position_span
                        .is_some_and(|[low, high]| low <= position && position <= high),
                    "position {position} on page {index}"
                );
            }
        }
        docx_layout::clear_measure_fonts();
    }

    #[test]
    fn an_edit_before_a_split_row_leaves_its_pages_unbuilt_with_moved_spans() {
        let (engine, extras) = split_row_engine(208);
        engine.set_display_window(Some(0..1));
        engine.set_windowed_incremental_builds(true);
        engine.build_display_list_frame(&extras, 0).unwrap();
        let before = engine.with_display_list(Clone::clone).unwrap();

        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 0),
                "Moved ",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        let incremental_builds = engine.stats().incremental_display_builds;
        let epoch = engine.display.borrow().binary_frame_epoch;
        engine.apply_and_layout("body", epoch).unwrap();
        assert_eq!(
            engine.stats().incremental_display_builds,
            incremental_builds + 1
        );
        let edited = engine.with_display_list(Clone::clone).unwrap();
        let windowed = {
            let pagination = engine.pagination.borrow();
            docx_layout::build_resident_display_list_partial_observed(
                pagination.input.as_ref().unwrap(),
                pagination.layout.as_ref().unwrap(),
                &extras,
                &|index| index == 0,
                &mut || {},
            )
            .unwrap()
            .1
        };
        assert_eq!(edited.pages, windowed.pages);
        let moved = edited.pages.iter().zip(&before.pages).skip(1);
        assert!(moved.len() >= 4);
        for (page, previous) in moved {
            let [low, high] = previous.position_span.unwrap();
            assert!(page.unbuilt);
            assert_eq!(page.position_span, Some([low + 6, high + 6]));
        }
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

    fn preview_mapped_snapshot(
        blocks: &[LayoutBlock],
        map: &LoweringMap,
        revealable: &[LayoutBlock],
    ) -> String {
        let spans: Vec<_> = map
            .spans
            .iter()
            .map(|span| {
                (
                    span.pm_start,
                    span.pm_end,
                    span.paragraph,
                    span.raw_start,
                    span.raw_end,
                    span.atom,
                )
            })
            .collect();
        serde_json::to_string(&(
            blocks,
            &map.stories,
            &map.paragraphs,
            &map.paragraph_blocks,
            spans,
            &map.tables,
            revealable,
        ))
        .unwrap()
    }

    fn preview_mapped_oracle(engine: &EngineSession, env: &RenderEnv) {
        engine.lower_story_json("body", env).unwrap();
        let actual = {
            let render = engine.render.borrow();
            let lowered = &render.stories["body"];
            preview_mapped_snapshot(&lowered.blocks, &lowered.map, &lowered.revealable_blocks)
        };
        let (blocks, map, revealable) =
            crate::bridge::yrs_doc_to_mapped_layout_blocks_with_revealable(
                engine.doc(),
                "body",
                env,
                &mut crate::bridge::local::LocalLowering::new(false),
            )
            .unwrap();
        assert_eq!(actual, preview_mapped_snapshot(&blocks, &map, &revealable));
    }

    fn preview_seeded(bytes: &[u8]) -> EngineSession {
        let engine = EngineSession::new(75210);
        crate::seed_from_docx(engine.doc(), bytes).unwrap();
        let primer = RenderEnv {
            show_hidden_text: true,
            ..RenderEnv::default()
        };
        engine.lower_story_json("body", &primer).unwrap();
        engine
    }

    #[test]
    fn preview_local_first_lowering_records_no_units() {
        let engine = EngineSession::new(75211);
        crate::seed_from_docx(engine.doc(), &preview_fixture::plain()).unwrap();
        let id = preview_fixture::ids(&engine).remove(0);
        preview_mapped_oracle(&engine, &RenderEnv::default());
        assert!(engine.render.borrow().stories["body"].preview.is_none());
        let before = engine.stats();
        preview_mapped_oracle(
            &engine,
            &RenderEnv::default().with_revision_preview(&id, RevisionPreview::Accepted),
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
        assert!(engine.render.borrow().stories["body"].preview.is_some());
        let before = engine.stats();
        preview_mapped_oracle(
            &engine,
            &RenderEnv::default().with_revision_preview(&id, RevisionPreview::Rejected),
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches + 1
        );
    }

    #[test]
    fn preview_local_mapped_oracle_decision_streams() {
        for (bytes, patches, fields) in [
            (preview_fixture::plain(), true, false),
            (preview_fixture::nested(), true, false),
            (preview_fixture::breaks(), false, false),
            (preview_fixture::drawings(), true, false),
            (preview_fixture::fields(), true, true),
            (preview_fixture::hidden_fields(), false, true),
            (preview_fixture::sequence(), false, true),
        ] {
            for seed in 0..40 {
                let engine = preview_seeded(&bytes);
                if fields {
                    preview_fixture::stamp_fields(&engine);
                }
                let ids = preview_fixture::ids(&engine);
                assert!(!ids.is_empty());
                let mut env = RenderEnv::default();
                let mut random = preview_fixture::Random::new(seed);
                preview_mapped_oracle(&engine, &env);
                for _ in 0..24 {
                    preview_fixture::decide(&mut env, &ids, &mut random);
                    let before = engine.stats();
                    preview_mapped_oracle(&engine, &env);
                    if patches {
                        assert_eq!(
                            engine.stats().lower_preview_patches,
                            before.lower_preview_patches + 1
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn preview_local_mapped_oracle_edit_streams() {
        for seed in 0..40 {
            let engine = preview_seeded(&preview_fixture::plain());
            let mut env = RenderEnv::default();
            let mut random = preview_fixture::Random::new(seed);
            preview_mapped_oracle(&engine, &env);
            for _ in 0..24 {
                let plain = crate::EditCtx::local("", "");
                let suggesting = crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting();
                let ctx = if random.next() % 2 == 0 {
                    &suggesting
                } else {
                    &plain
                };
                if random.next() % 2 == 0 {
                    engine
                        .doc()
                        .insert_text(
                            ctx,
                            crate::Position::new("body", 1),
                            "x",
                            crate::FormatPolicy::Plain,
                        )
                        .unwrap();
                } else {
                    engine
                        .doc()
                        .delete_range(ctx, crate::StoryRange::new("body", 1, 2))
                        .unwrap();
                }
                preview_mapped_oracle(&engine, &env);
                let ids = preview_fixture::ids(&engine);
                preview_fixture::decide(&mut env, &ids, &mut random);
                let before = engine.stats();
                preview_mapped_oracle(&engine, &env);
                assert_eq!(
                    engine.stats().lower_preview_patches,
                    before.lower_preview_patches + 1
                );
            }
        }
    }

    #[test]
    fn preview_local_mapped_oracle_corpus() {
        for (name, bytes) in preview_fixture::corpus() {
            let engine = preview_seeded(bytes);
            let ids = preview_fixture::ids(&engine);
            assert!(!ids.is_empty(), "{name}");
            preview_mapped_oracle(&engine, &RenderEnv::default());
            for id in &ids {
                for decision in [RevisionPreview::Accepted, RevisionPreview::Rejected] {
                    preview_mapped_oracle(
                        &engine,
                        &RenderEnv::default().with_revision_preview(id, decision),
                    );
                }
            }
            for seed in 0..40 {
                let mut env = RenderEnv::default();
                let mut random = preview_fixture::Random::new(seed);
                for _ in 0..12 {
                    preview_fixture::decide(&mut env, &ids, &mut random);
                    preview_mapped_oracle(&engine, &env);
                }
            }
        }
    }

    #[test]
    fn preview_local_preserves_shared_snapshots() {
        let engine = preview_seeded(&preview_fixture::drawings());
        let env = RenderEnv::default();
        preview_mapped_oracle(&engine, &env);
        let (blocks, map, revealable) = {
            let render = engine.render.borrow();
            let lowered = &render.stories["body"];
            (
                lowered.blocks.clone(),
                lowered.map.clone(),
                lowered.revealable_blocks.clone(),
            )
        };
        let original = preview_mapped_snapshot(&blocks, &map, &revealable);
        let env = env.with_revision_preview("1", RevisionPreview::Rejected);
        let before = engine.stats();
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches + 1
        );
        assert_eq!(
            original,
            preview_mapped_snapshot(&blocks, &map, &revealable)
        );
        let render = engine.render.borrow();
        let lowered = &render.stories["body"];
        assert!(!Rc::ptr_eq(&blocks, &lowered.blocks));
        assert!(!Rc::ptr_eq(&map, &lowered.map));
        assert!(!Rc::ptr_eq(&revealable, &lowered.revealable_blocks));
    }

    #[test]
    fn preview_local_no_target_preserves_cache_values() {
        let engine = preview_seeded(&preview_fixture::plain());
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let (blocks, map, revealable, serialized, units) = {
            let render = engine.render.borrow();
            let lowered = &render.stories["body"];
            (
                lowered.blocks.clone(),
                lowered.map.clone(),
                lowered.revealable_blocks.clone(),
                lowered.serialized_blocks.clone(),
                lowered.preview.clone().unwrap(),
            )
        };
        let env = RenderEnv::default().with_revision_preview("unused", RevisionPreview::Rejected);
        let before = engine.stats();
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches + 1
        );
        assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
        let render = engine.render.borrow();
        let lowered = &render.stories["body"];
        assert!(Rc::ptr_eq(&blocks, &lowered.blocks));
        assert!(Rc::ptr_eq(&map, &lowered.map));
        assert!(Rc::ptr_eq(&revealable, &lowered.revealable_blocks));
        assert!(Rc::ptr_eq(&units, lowered.preview.as_ref().unwrap()));
        assert_eq!(serialized, lowered.serialized_blocks);
        assert_eq!(lowered.env, env);
    }

    #[test]
    fn preview_local_live_text_seeds_fall_back() {
        let bytes = preview_fixture::document(&preview_fixture::paragraph(
            1,
            &preview_fixture::run("Plain"),
        ));
        let engine = preview_seeded(&bytes);
        engine.set_local_lowering(true);
        preview_mapped_oracle(&engine, &RenderEnv::default());
        assert!(!engine.render.borrow().stories["body"].local.blocked);
        let before = engine.stats();
        preview_mapped_oracle(
            &engine,
            &RenderEnv::default().with_revision_preview("unused", RevisionPreview::Accepted),
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
        assert_eq!(
            engine.stats().lower_preview_fallbacks,
            before.lower_preview_fallbacks + 1
        );
        assert!(!engine.render.borrow().stories["body"].local.blocked);
    }

    #[test]
    fn preview_local_after_state_mismatch_discards_all_targets() {
        let numbered = preview_fixture::paragraph(
            1,
            r#"<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Cell</w:t></w:r>"#,
        );
        let bytes = preview_fixture::document(&format!(
            "{}{}{}",
            preview_fixture::paragraph(
                2,
                &preview_fixture::revision("ins", "1", &preview_fixture::run("First"))
            ),
            preview_fixture::table(&numbered),
            preview_fixture::paragraph(3, &preview_fixture::run("After"))
        ));
        let engine = preview_seeded(&bytes);
        let stamp = Any::Map(std::sync::Arc::new(HashMap::from([(
            "id".to_owned(),
            Any::from("table"),
        )])));
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![crate::RawOp::Format {
                    index: 6,
                    len: 1,
                    attrs: [("ins".into(), stamp)].into(),
                }],
                &crate::EditCtx::local("", ""),
            )
            .unwrap();
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let old = engine.render.borrow().stories["body"].blocks.clone();
        let original = serde_json::to_string(&old).unwrap();
        let env = RenderEnv::default()
            .with_revision_preview("1", RevisionPreview::Rejected)
            .with_revision_preview("table", RevisionPreview::Rejected);
        let before = engine.stats();
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
        assert_eq!(
            engine.stats().lower_preview_fallbacks,
            before.lower_preview_fallbacks + 1
        );
        assert_eq!(original, serde_json::to_string(&old).unwrap());
    }

    #[test]
    fn preview_local_leading_tracked_break_falls_back() {
        let engine = preview_seeded(&preview_fixture::document(&preview_fixture::paragraph(
            1,
            &preview_fixture::run("After"),
        )));
        let receipt = engine
            .doc()
            .insert_embed(
                &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                crate::Position::new("body", 0),
                "pageBreak",
                vec![],
            )
            .unwrap();
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let before = engine.stats();
        let env = RenderEnv::default()
            .with_revision_preview(&receipt.revision_ids[0], RevisionPreview::Rejected);
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
        assert_eq!(
            engine.stats().lower_preview_fallbacks,
            before.lower_preview_fallbacks + 1
        );
    }

    #[test]
    fn preview_local_epoch_and_environment_changes_lower_fully() {
        let engine = preview_seeded(&preview_fixture::plain());
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let mut env = RenderEnv::default().with_revision_preview("1", RevisionPreview::Accepted);
        env.show_hidden_text = true;
        let before = engine.stats();
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_cache_misses,
            before.lower_cache_misses + 1
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", 1),
                "x",
                crate::FormatPolicy::Plain,
            )
            .unwrap();
        env.revision_preview
            .insert("1".to_owned(), RevisionPreview::Rejected);
        let before = engine.stats();
        preview_mapped_oracle(&engine, &env);
        assert_eq!(
            engine.stats().lower_cache_misses,
            before.lower_cache_misses + 1
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches
        );
    }

    #[test]
    fn preview_local_region_layout_matches_fresh_engine() {
        let bytes = preview_fixture::breaks();
        let engine = preview_seeded(&bytes);
        let font = docx_layout::register_measure_font(LIBERATION).unwrap();
        let mut request: serde_json::Value =
            serde_json::from_str(&small_page_request(font)).unwrap();
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        request["renderEnv"] = json!({"revisionPreview": {"1": "rejected"}});
        let before = engine.stats();
        let actual = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches + 1
        );
        let fresh = preview_seeded(&bytes);
        let expected = fresh
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        assert_eq!(actual, expected);
    }
    #[test]
    fn preview_local_sequence_marker_without_fields_patches() {
        use yrs::{Map, ReadTxn};
        let engine = preview_seeded(&preview_fixture::plain());
        let mut txn = engine.doc().transact_for(&crate::EditCtx::system(""));
        let session = txn.get_map(crate::identity::SESSION).unwrap();
        session.insert(
            &mut txn,
            crate::seed::OPAQUE_SEQUENCES,
            Any::Array(Vec::new().into()),
        );
        drop(txn);
        preview_mapped_oracle(&engine, &RenderEnv::default());
        assert!(engine.render.borrow().stories["body"].preview.is_some());
        let before = engine.stats();
        preview_mapped_oracle(
            &engine,
            &RenderEnv::default().with_revision_preview("1", RevisionPreview::Rejected),
        );
        assert_eq!(
            engine.stats().lower_preview_patches,
            before.lower_preview_patches + 1
        );
    }

    #[test]
    fn preview_local_revision_attrs_block_text_seeds() {
        for imported in [false, true] {
            let bytes = if imported {
                preview_fixture::plain()
            } else {
                preview_fixture::document(&preview_fixture::paragraph(
                    1,
                    &preview_fixture::run("Plain"),
                ))
            };
            let engine = preview_seeded(&bytes);
            engine.set_local_lowering(true);
            let id = if imported {
                "1".to_owned()
            } else {
                engine
                    .doc()
                    .insert_text(
                        &crate::EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                        crate::Position::new("body", 1),
                        "x",
                        crate::FormatPolicy::Plain,
                    )
                    .unwrap()
                    .revision_ids[0]
                    .clone()
            };
            preview_mapped_oracle(&engine, &RenderEnv::default());
            assert!(engine.render.borrow().stories["body"].local.blocked);
            let before = engine.stats();
            preview_mapped_oracle(
                &engine,
                &RenderEnv::default().with_revision_preview(id, RevisionPreview::Rejected),
            );
            assert_eq!(
                engine.stats().lower_preview_patches,
                before.lower_preview_patches + 1
            );
        }
    }
}
