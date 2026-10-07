use std::cell::RefCell;
use std::collections::btree_map::Entry;
use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_text::measure::{FontChainDependencies, FontChains};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::cell_layout::{nested_table_float_offset, nested_table_horizontal_offset};
use crate::floating_objects::{MIN_WRAP_SEGMENT_WIDTH, table_wrap_gaps};
use crate::table_grid::{
    ResolvedGridCell, content_sized_columns, count_table_columns, fits_columns_to_words,
    grow_content_sized_columns, resolve_cell_grid,
    resolve_table_column_widths_with_percentage_basis, resolve_table_width_px,
    table_percentage_basis, widen_columns_to_minimums,
};
use crate::types::{
    BlockExtent, BlockId, ChartExtent, FloatingTablePosition, ImageExtent, ImageRunPosition,
    LayoutBlock, ParagraphBlock, ParagraphExtent, ParagraphSpacing, Run, ShapeBlock, ShapeExtent,
    TableBlock, TableCellExtent, TableExtent, TableRowExtent, TextBoxBlock, TextBoxExtent,
    TypesetBidiSlice, TypesetClusterAdvance, TypesetRow, TypesetRowSegment, TypesetRunAdvance,
};
use ooxml_text::{FontSlotUse, LineBox, LineSpacingRule, apply_spacing_rule, font_slot_use};

const DEFAULT_CELL_PADDING_X: f64 = 7.0;
const DEFAULT_CELL_PADDING_Y: f64 = 0.0;
/// Zones one anchor frame may accumulate, matching the measurement layer's cap.
const MAX_ACTIVE_ZONES: usize = 200;

/// A shape the lowering placed by anchor rather than in the flow.
fn anchored_shape(shape: &ShapeBlock) -> bool {
    shape.position.is_some() || shape.wrap_type.is_some()
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct FloatStrip {
    pub(crate) left_offset: f64,
    pub(crate) available_width: f64,
}

#[derive(Clone, Debug)]
pub(crate) struct FloatingZone {
    pub(crate) left_margin: f64,
    pub(crate) right_margin: f64,
    pub(crate) top_y: f64,
    pub(crate) bottom_y: f64,
    /// Usable strips when the float sits inside the text column and text runs
    /// past it on both sides; empty means the side margins describe the zone.
    pub(crate) segments: Vec<FloatStrip>,
    pub(crate) full_width_block: bool,
}

#[derive(Clone, Debug)]
struct AnchoredFloatingZone {
    zone: FloatingZone,
    anchor_block_index: usize,
    margin_relative: bool,
}

#[derive(Clone, Debug)]
pub struct FloatPageGeometry {
    pub page_width: f64,
    pub margin_left: f64,
    pub page_height: f64,
    pub margin_top: f64,
    pub content_height: f64,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MeasurementConfig {
    #[serde(default)]
    pub font_chains: BTreeMap<String, Vec<u32>>,
    #[serde(default)]
    pub defaults: Value,
    #[serde(default)]
    pub compat: Value,
    #[serde(default = "default_true")]
    pub authoritative_shaping: bool,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FontRequirement {
    pub key: String,
    pub family: String,
    pub bold: bool,
    pub italic: bool,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub scripts: Vec<String>,
}

/// The fonts `blocks` need, text naming no family taking `default_family`, the family
/// measurement defaults to.
pub fn collect_font_requirements<'a>(
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
    default_family: &str,
) -> Vec<FontRequirement> {
    let mut collector = FontRequirementCollector::default();
    collector.collect(blocks, default_family);
    collector.finish().into_values().collect()
}

/// [`collect_font_requirements`] into `requirements`, keyed as it keys them,
/// for callers gathering several block runs without copying them. It keeps
/// every family the blocks name and only those, since a later call may reach
/// one; gather with a [`FontRequirementCollector`] to drop the unused ones.
pub fn collect_font_requirements_into<'a>(
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
    default_family: &str,
    requirements: &mut BTreeMap<String, FontRequirement>,
) {
    let mut collector = FontRequirementCollector::default();
    collector.collect(blocks, default_family);
    collector.merge_named_into(requirements);
}

/// Fonts any revision preview may need. Returns false when any of them needs a
/// script fallback, whose chain order a superset cannot keep exact. Keeps every
/// named family, as [`collect_font_requirements_into`] does.
pub fn collect_preview_font_requirements_into<'a>(
    blocks: impl IntoIterator<Item = &'a LayoutBlock>,
    default_family: &str,
    requirements: &mut BTreeMap<String, FontRequirement>,
) -> bool {
    let mut collector = FontRequirementCollector::default();
    collector.collect_preview(blocks, default_family);
    collector.merge_named_into(requirements);
    requirements
        .values()
        .all(|requirement| requirement.scripts.is_empty())
}

/// Gathers the fonts of several block runs (a document's body, headers, footers
/// and notes) and keeps, once all are in, those measurement can reach, each with
/// the scripts of every run naming it in order, so dropping an unused slot keeps
/// fallback order.
#[derive(Default)]
pub struct FontRequirementCollector {
    named: BTreeMap<String, FontRequirement>,
    implicit: BTreeMap<String, FontRequirement>,
    used: HashSet<String>,
}

impl FontRequirementCollector {
    /// Adds the fonts `blocks` need, as [`collect_font_requirements`] does.
    pub fn collect<'a>(
        &mut self,
        blocks: impl IntoIterator<Item = &'a LayoutBlock>,
        default_family: &str,
    ) {
        for block in blocks {
            walk_paragraphs(std::slice::from_ref(block), &mut |paragraph| {
                let scripts = paragraph_scripts(paragraph);
                collect_paragraph_font_requirements(paragraph, &scripts, default_family, self);
            });
        }
    }

    /// Adds the fonts any revision preview of `blocks` may need, as
    /// [`collect_preview_font_requirements_into`] does.
    pub fn collect_preview<'a>(
        &mut self,
        blocks: impl IntoIterator<Item = &'a LayoutBlock>,
        default_family: &str,
    ) {
        let blocks: Vec<&LayoutBlock> = blocks.into_iter().collect();
        // A preview hiding the drawings that split a paragraph joins its segments.
        let mut segment_scripts = HashMap::<String, Vec<String>>::new();
        for block in &blocks {
            walk_paragraphs(std::slice::from_ref(*block), &mut |paragraph| {
                if let BlockId::Str(id) = &paragraph.id
                    && !id.is_empty()
                {
                    let scripts = segment_scripts.entry(id.clone()).or_default();
                    for script in paragraph_scripts_with_han_fallback(paragraph, true) {
                        if !scripts.contains(&script) {
                            scripts.push(script);
                        }
                    }
                }
            });
        }
        for block in blocks {
            walk_paragraphs(std::slice::from_ref(block), &mut |paragraph| {
                let scripts = match &paragraph.id {
                    BlockId::Str(id) if !id.is_empty() => segment_scripts[id].clone(),
                    _ => paragraph_scripts_with_han_fallback(paragraph, true),
                };
                collect_paragraph_font_requirements(paragraph, &scripts, default_family, self);
                let Some(attrs) = &paragraph.attrs else {
                    return;
                };
                if attrs.num_pr.is_none()
                    && attrs.list_marker.is_none()
                    && attrs.list_is_bullet.is_none()
                    && attrs.list_marker_hidden.is_none()
                    && attrs.list_marker_font_family.is_none()
                    && attrs.list_marker_font_size.is_none()
                    && attrs.list_marker_bold.is_none()
                    && attrs.list_marker_italic.is_none()
                    && attrs.list_marker_color.is_none()
                    && attrs.list_marker_suffix.is_none()
                    && attrs.list_marker_revision.is_none()
                {
                    return;
                }
                let default_family = attrs
                    .default_font_family
                    .as_deref()
                    .unwrap_or(default_family);
                let marker_style = (
                    attrs.list_marker_bold.unwrap_or(false),
                    attrs.list_marker_italic.unwrap_or(false),
                );
                let styles = if attrs.list_marker.is_none() {
                    &[(false, false), (true, false), (false, true), (true, true)][..]
                } else {
                    std::slice::from_ref(&marker_style)
                };
                for family in paragraph
                    .runs
                    .iter()
                    .filter_map(|run| {
                        let formatting = match run {
                            Run::Text(text) => &text.fmt,
                            Run::Tab(tab) => &tab.fmt,
                            Run::Field(field) => &field.fmt,
                            _ => return None,
                        };
                        formatting.font_family.as_deref()
                    })
                    .chain(std::iter::once(default_family))
                {
                    let family = attrs.list_marker_font_family.as_deref().unwrap_or(family);
                    for &(bold, italic) in styles {
                        self.name(family, bold, italic, &scripts, true);
                    }
                }
            });
        }
    }

    /// The requirements measurement can reach, keyed as
    /// [`collect_font_requirements`] keys them.
    pub fn finish(self) -> BTreeMap<String, FontRequirement> {
        let mut requirements = BTreeMap::new();
        self.merge_into(&mut requirements);
        requirements
    }

    /// Merges every named requirement and nothing else, as main collects them:
    /// an implicit entry kept here would lead the scripts a later call names.
    fn merge_named_into(mut self, requirements: &mut BTreeMap<String, FontRequirement>) {
        self.used.extend(self.named.keys().cloned());
        self.implicit.clear();
        self.merge_into(requirements);
    }

    fn merge_into(self, requirements: &mut BTreeMap<String, FontRequirement>) {
        let Self {
            named,
            implicit,
            used,
        } = self;
        for (key, requirement) in implicit {
            if !named.contains_key(&key) {
                requirements.entry(key).or_insert(requirement);
            }
        }
        for (key, requirement) in named {
            match requirements.entry(key) {
                Entry::Occupied(mut kept) => {
                    let scripts = &mut kept.get_mut().scripts;
                    for script in requirement.scripts {
                        if !scripts.contains(&script) {
                            scripts.push(script);
                        }
                    }
                }
                Entry::Vacant(slot) if used.contains(slot.key()) => {
                    slot.insert(requirement);
                }
                Entry::Vacant(_) => {}
            }
        }
    }

    /// A family and style the text names, which measurement reaches when `reached`.
    fn name(&mut self, family: &str, bold: bool, italic: bool, scripts: &[String], reached: bool) {
        let key = add_font_requirement(family, bold, italic, scripts, &mut self.named);
        if reached {
            self.used.insert(key);
        }
    }

    /// A family and style measurement resolves a slot the text leaves unnamed to.
    fn reach(&mut self, family: &str, bold: bool, italic: bool, scripts: &[String]) {
        let key = add_font_requirement(family, bold, italic, scripts, &mut self.implicit);
        self.used.insert(key);
    }
}

/// The family measurement gives text naming none: `defaults.fontFamily`, else Calibri.
pub fn default_font_family(defaults: &Value) -> &str {
    defaults
        .get("fontFamily")
        .and_then(Value::as_str)
        .unwrap_or("Calibri")
}

/// Whether any line of `extents` carries synthetic metrics because measuring it failed.
pub fn measured_synthetically<'a>(extents: impl IntoIterator<Item = &'a BlockExtent>) -> bool {
    fn synthetic(lines: &[crate::types::TypesetRow]) -> bool {
        lines
            .iter()
            .any(|line| line.synthetic_fallback == Some(true))
    }
    fn block(extent: &BlockExtent) -> bool {
        match extent {
            BlockExtent::Paragraph(paragraph) => synthetic(&paragraph.lines),
            BlockExtent::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .flat_map(|cell| &cell.blocks)
                .any(block),
            BlockExtent::TextBox(text_box) => text_box
                .inner_measures
                .iter()
                .any(|paragraph| synthetic(&paragraph.lines)),
            _ => false,
        }
    }
    extents.into_iter().any(block)
}

fn walk_paragraphs(blocks: &[LayoutBlock], visit: &mut impl FnMut(&ParagraphBlock)) {
    for block in blocks {
        match block {
            LayoutBlock::Paragraph(paragraph) => visit(paragraph),
            LayoutBlock::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        walk_paragraphs(&cell.blocks, visit);
                    }
                }
            }
            LayoutBlock::TextBox(text_box) => {
                for paragraph in &text_box.content {
                    visit(paragraph);
                }
            }
            LayoutBlock::Shape(shape) => {
                if let Some(paragraphs) = &shape.inner_text {
                    for paragraph in paragraphs {
                        visit(paragraph);
                    }
                }
                for child in &shape.children {
                    walk_shape_paragraphs(child, visit);
                }
            }
            _ => {}
        }
    }
}

fn walk_shape_paragraphs(shape: &ShapeBlock, visit: &mut impl FnMut(&ParagraphBlock)) {
    if let Some(paragraphs) = &shape.inner_text {
        for paragraph in paragraphs {
            visit(paragraph);
        }
    }
    for child in &shape.children {
        walk_shape_paragraphs(child, visit);
    }
}

fn add_font_requirement(
    family: &str,
    bold: bool,
    italic: bool,
    scripts: &[String],
    requirements: &mut BTreeMap<String, FontRequirement>,
) -> String {
    let key = format!(
        "{}|{}|{}",
        family.to_lowercase(),
        u8::from(bold),
        u8::from(italic)
    );
    let requirement = requirements
        .entry(key.clone())
        .or_insert_with(|| FontRequirement {
            key: key.clone(),
            family: family.to_owned(),
            bold,
            italic,
            scripts: Vec::new(),
        });
    for script in scripts {
        if !requirement.scripts.contains(script) {
            requirement.scripts.push(script.clone());
        }
    }
    key
}

/// Names every family and style `paragraph` names and reaches those measurement
/// takes an unnamed slot from.
fn collect_paragraph_font_requirements(
    paragraph: &ParagraphBlock,
    scripts: &[String],
    fallback_family: &str,
    collector: &mut FontRequirementCollector,
) {
    let default_family = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.default_font_family.as_deref())
        .unwrap_or(fallback_family);
    collector.name(default_family, false, false, scripts, true);
    for run in &paragraph.runs {
        let (formatting, text) = match run {
            Run::Text(text) => (&text.fmt, Some(text)),
            Run::Tab(tab) => (&tab.fmt, None),
            Run::Field(field) => (&field.fmt, None),
            _ => continue,
        };
        let include_regular = text.is_some();
        let bold = formatting.bold.unwrap_or(false);
        let italic = formatting.italic.unwrap_or(false);
        let slots = formatting.font_slots.as_ref();
        let complex_script = formatting.complex_script.unwrap_or(false);
        let mut slot_use = text.map_or_else(FontSlotUse::default, |text| {
            font_slot_use(
                &text.text,
                complex_script,
                slots.and_then(|slots| slots.hint.as_deref()),
            )
        });
        // Text typed into a complex-script run measures with its cs slot too.
        slot_use.complex_script |= complex_script;
        let cs_bold = formatting.bold_cs.unwrap_or(bold);
        let cs_italic = formatting.italic_cs.unwrap_or(italic);
        let mut name = |family: &str, bold: bool, italic: bool, reached: bool| {
            collector.name(family, bold, italic, scripts, reached);
            if include_regular {
                collector.name(family, false, false, scripts, reached);
            }
        };
        name(
            formatting.font_family.as_deref().unwrap_or(default_family),
            bold,
            italic,
            true,
        );
        if let Some(slots) = slots {
            for (family, bold, italic, reached) in [
                (slots.ascii.as_deref(), bold, italic, true),
                (slots.h_ansi.as_deref(), bold, italic, true),
                (slots.east_asia.as_deref(), bold, italic, slot_use.east_asia),
                (
                    slots.cs.as_deref(),
                    cs_bold,
                    cs_italic,
                    slot_use.complex_script,
                ),
            ] {
                if let Some(family) = family {
                    name(family, bold, italic, reached);
                }
            }
        }
        // What measurement takes a slot the run leaves unnamed from
        // (`family_for_slot` in ooxml-text).
        let run_family = formatting.font_family.as_deref().unwrap_or(fallback_family);
        let text_unnamed = slots
            .and_then(|slots| slots.h_ansi.as_deref().or(slots.ascii.as_deref()))
            .unwrap_or(run_family);
        let unnamed = if text.is_some() {
            text_unnamed
        } else {
            run_family
        };
        collector.reach(unnamed, bold, italic, scripts);
        // An empty text run takes its metrics from the run's family alone.
        collector.reach(run_family, bold, italic, scripts);
        if slot_use.east_asia {
            collector.reach(
                slots
                    .and_then(|slots| slots.east_asia.as_deref())
                    .unwrap_or(unnamed),
                bold,
                italic,
                scripts,
            );
        }
        // Text typed into a tab's or field's complex-script run resolves as text does.
        if slot_use.complex_script {
            collector.reach(
                slots
                    .and_then(|slots| slots.cs.as_deref())
                    .unwrap_or(text_unnamed),
                cs_bold,
                cs_italic,
                scripts,
            );
        }
    }
    if let Some(attrs) = &paragraph.attrs
        && attrs
            .list_marker
            .as_deref()
            .is_some_and(|marker| !marker.is_empty())
        && attrs.list_marker_hidden != Some(true)
    {
        let first_run_family = paragraph.runs.iter().find_map(|run| match run {
            Run::Text(text) => text.fmt.font_family.as_deref(),
            _ => None,
        });
        collector.name(
            attrs
                .list_marker_font_family
                .as_deref()
                .or(first_run_family)
                .unwrap_or(default_family),
            attrs.list_marker_bold.unwrap_or(false),
            attrs.list_marker_italic.unwrap_or(false),
            scripts,
            false,
        );
        let first_text_run = paragraph.runs.iter().find_map(|run| match run {
            Run::Text(text) => Some(text),
            _ => None,
        });
        collector.reach(
            attrs
                .list_marker_font_family
                .as_deref()
                .or_else(|| first_text_run.and_then(|text| text.fmt.font_family.as_deref()))
                .unwrap_or(default_family),
            attrs.list_marker_bold.unwrap_or(false),
            attrs.list_marker_italic.unwrap_or(false),
            scripts,
        );
    }
}

fn paragraph_scripts(paragraph: &ParagraphBlock) -> Vec<String> {
    paragraph_scripts_with_han_fallback(paragraph, false)
}

fn paragraph_scripts_with_han_fallback(
    paragraph: &ParagraphBlock,
    include_han_fallback: bool,
) -> Vec<String> {
    let mut han = false;
    let mut kana = false;
    let mut hangul = false;
    let mut arabic = false;
    let mut hebrew = false;
    for text in paragraph.runs.iter().filter_map(|run| match run {
        Run::Text(text) => Some(text.text.as_str()),
        _ => None,
    }) {
        for character in text.chars() {
            let point = character as u32;
            match point {
                0x0590..=0x05ff | 0xfb1d..=0xfb4f => hebrew = true,
                0x0600..=0x06ff
                | 0x0750..=0x077f
                | 0x0870..=0x08ff
                | 0xfb50..=0xfdff
                | 0xfe70..=0xfeff => arabic = true,
                0x1100..=0x11ff
                | 0x3130..=0x318f
                | 0xa960..=0xa97f
                | 0xac00..=0xd7ff
                | 0xffa0..=0xffdc => hangul = true,
                0x3040..=0x30ff | 0x31f0..=0x31ff | 0xff66..=0xff9f => kana = true,
                0x3000..=0x303f
                | 0x3400..=0x4dbf
                | 0x4e00..=0x9fff
                | 0xf900..=0xfaff
                | 0xfe30..=0xfe4f
                | 0xff00..=0xff65
                | 0x20000..=0x3ffff => han = true,
                _ => {}
            }
        }
    }
    let mut scripts = Vec::new();
    if kana {
        scripts.push("cjk-jp".to_owned());
    }
    if hangul {
        scripts.push("cjk-kr".to_owned());
    }
    if han && (include_han_fallback || (!kana && !hangul)) {
        scripts.push("cjk-sc".to_owned());
    }
    if arabic {
        scripts.push("arabic".to_owned());
    }
    if hebrew {
        scripts.push("hebrew".to_owned());
    }
    scripts
}

fn default_true() -> bool {
    true
}

pub fn measure_blocks(
    blocks: &mut [LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<Vec<BlockExtent>, String> {
    blocks
        .iter_mut()
        .map(|block| measure_block(block, content_width, config))
        .collect()
}

pub fn measure_blocks_without_table_compat_shift(
    blocks: &mut [LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<Vec<BlockExtent>, String> {
    blocks
        .iter_mut()
        .map(|block| match block {
            LayoutBlock::Table(table) => {
                measure_table_with_compat_shift(table, content_width, config, false)
                    .map(BlockExtent::Table)
            }
            _ => measure_block(block, content_width, config),
        })
        .collect()
}

/// Whether any block anchors a floating zone (wrapped image, floating table,
/// or text box). Callers use this to gate float-free fast paths; extraction is
/// a read-only scan of the same zones `measure_blocks_with_floats` consumes.
pub fn has_floating_zones(
    blocks: &[LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
) -> Result<bool, String> {
    Ok(!extract_floating_zones(
        blocks,
        content_width,
        &[],
        &[],
        config,
        page_geometry,
        &BTreeMap::new(),
    )?
    .is_empty())
}

pub fn measure_blocks_with_floats(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
) -> Result<Vec<BlockExtent>, String> {
    measure_blocks_with_shape_offsets(blocks, widths, config, page_geometry, &BTreeMap::new())
}

pub fn measure_blocks_with_shape_offsets(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    shape_offsets: &BTreeMap<usize, f64>,
) -> Result<Vec<BlockExtent>, String> {
    measure_blocks_with_table_wrap_frames(blocks, widths, &[], config, page_geometry, shape_offsets)
}

/// Float measurement with known single-column table frames.
pub fn measure_blocks_with_table_wrap_frames(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    table_wrap_frames: &[bool],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    shape_offsets: &BTreeMap<usize, f64>,
) -> Result<Vec<BlockExtent>, String> {
    let default_width = widths.first().copied().unwrap_or(0.0);
    let extracted = extract_floating_zones(
        blocks,
        default_width,
        widths,
        table_wrap_frames,
        config,
        page_geometry,
        shape_offsets,
    )?;
    let (paragraph_zones, zones_by_anchor) = group_floating_zones(extracted);
    let marks = section_break_marks(blocks);
    measure_float_flow(
        blocks,
        widths,
        default_width,
        config,
        &paragraph_zones,
        &zones_by_anchor,
        &marks,
    )
}

/// Whether float measurement clears its floating zones at `block`: page,
/// column and section breaks and a paragraph that breaks the page before it.
pub fn resets_float_flow(block: &LayoutBlock) -> bool {
    matches!(
        block,
        LayoutBlock::PageBreak(_) | LayoutBlock::ColumnBreak(_) | LayoutBlock::SectionBreak(_)
    ) || crate::keep_together::paragraph_breaks_before(block)
}

/// `(anchors a floating zone, anchors a margin-relative one)` for `blocks`.
/// Margin-relative zones are shared across the document; paragraph-relative
/// ones never reach past the next [`resets_float_flow`] block.
pub fn floating_zone_kinds(
    blocks: &[LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
) -> Result<(bool, bool), String> {
    let zones = extract_floating_zones(
        blocks,
        content_width,
        &[],
        &[],
        config,
        page_geometry,
        &BTreeMap::new(),
    )?;
    Ok((
        !zones.is_empty(),
        zones.iter().any(|zone| zone.margin_relative),
    ))
}

/// Float-aware extents for one flow segment, equal to what
/// [`measure_blocks_with_floats`] gives those blocks within their document
/// when it has no margin-relative zones. The segment starts at the document
/// start or a [`resets_float_flow`] block and stops before the next one;
/// `default_width` and `section_break_marks` come from the whole document.
/// `Ok(None)` when the segment anchors a margin-relative zone.
pub fn measure_float_segment(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    default_width: f64,
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    section_break_marks: &[bool],
) -> Result<Option<Vec<BlockExtent>>, String> {
    measure_float_segment_with_table_wrap_frames(
        blocks,
        widths,
        default_width,
        &[],
        config,
        page_geometry,
        section_break_marks,
    )
}

/// Segment measurement with known single-column table frames.
pub fn measure_float_segment_with_table_wrap_frames(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    default_width: f64,
    table_wrap_frames: &[bool],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    section_break_marks: &[bool],
) -> Result<Option<Vec<BlockExtent>>, String> {
    measure_float_segment_with_font_dependencies(
        blocks,
        widths,
        default_width,
        table_wrap_frames,
        config,
        page_geometry,
        section_break_marks,
    )
    .map(|measured| measured.map(|(extents, _)| extents))
}

pub fn measure_float_segment_with_font_dependencies(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    default_width: f64,
    table_wrap_frames: &[bool],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    section_break_marks: &[bool],
) -> Result<Option<(Vec<BlockExtent>, Vec<FontChainDependencies>)>, String> {
    let (extracted, zone_dependencies) = extract_floating_zones_recorded(
        blocks,
        default_width,
        widths,
        table_wrap_frames,
        config,
        page_geometry,
        &BTreeMap::new(),
        true,
    )?;
    if extracted.iter().any(|zone| zone.margin_relative) {
        return Ok(None);
    }
    let (paragraph_zones, zones_by_anchor) = group_floating_zones(extracted);
    measure_float_flow_recorded(
        blocks,
        widths,
        default_width,
        config,
        &paragraph_zones,
        &zones_by_anchor,
        section_break_marks,
        &zone_dependencies,
    )
    .map(Some)
}

type ParagraphZones = BTreeMap<usize, Vec<FloatingZone>>;
type AnchorZones = HashMap<usize, Vec<FloatingZone>>;

fn group_floating_zones(extracted: Vec<AnchoredFloatingZone>) -> (ParagraphZones, AnchorZones) {
    let mut margin_groups = BTreeMap::<u64, Vec<AnchoredFloatingZone>>::new();
    let mut paragraph_zones = ParagraphZones::new();
    for anchored in extracted {
        if anchored.margin_relative {
            margin_groups
                .entry(anchored.zone.top_y.to_bits())
                .or_default()
                .push(anchored);
        } else {
            paragraph_zones
                .entry(anchored.anchor_block_index)
                .or_default()
                .push(anchored.zone);
        }
    }
    let mut zones_by_anchor = AnchorZones::new();
    for group in margin_groups.into_values() {
        let earliest = group
            .iter()
            .map(|anchored| anchored.anchor_block_index)
            .min()
            .unwrap_or(0);
        for anchored in group {
            let anchor = if anchored.zone.full_width_block {
                0
            } else {
                earliest
            };
            zones_by_anchor
                .entry(anchor)
                .or_default()
                .push(anchored.zone);
        }
    }
    (paragraph_zones, zones_by_anchor)
}

/// Bare paragraph marks that carry a section break and print no line.
pub fn section_break_marks(blocks: &[LayoutBlock]) -> Vec<bool> {
    blocks
        .iter()
        .enumerate()
        .map(|(index, block)| {
            is_section_break_mark(
                block,
                index
                    .checked_sub(1)
                    .and_then(|previous| blocks.get(previous)),
                blocks.get(index + 1),
            )
        })
        .collect()
}

pub fn is_section_break_mark(
    block: &LayoutBlock,
    previous: Option<&LayoutBlock>,
    next: Option<&LayoutBlock>,
) -> bool {
    matches!(block, LayoutBlock::Paragraph(paragraph) if paragraph.runs.is_empty())
        && matches!(next, Some(LayoutBlock::SectionBreak(_)))
        && previous.is_some_and(|block| !matches!(block, LayoutBlock::SectionBreak(_)))
}

fn measure_float_flow(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    default_width: f64,
    config: &MeasurementConfig,
    paragraph_zones: &ParagraphZones,
    zones_by_anchor: &AnchorZones,
    section_break_marks: &[bool],
) -> Result<Vec<BlockExtent>, String> {
    let mut flow = FlowState::default();
    let mut measured = Vec::with_capacity(blocks.len());
    for (index, block) in blocks.iter_mut().enumerate() {
        measured.push(flow.measure(
            index,
            block,
            widths,
            default_width,
            config,
            paragraph_zones,
            zones_by_anchor,
            section_break_marks,
        )?);
    }
    Ok(measured)
}

#[allow(clippy::too_many_arguments)]
fn measure_float_flow_recorded(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    default_width: f64,
    config: &MeasurementConfig,
    paragraph_zones: &ParagraphZones,
    zones_by_anchor: &AnchorZones,
    section_break_marks: &[bool],
    zone_dependencies: &[FontChainDependencies],
) -> Result<(Vec<BlockExtent>, Vec<FontChainDependencies>), String> {
    let mut flow = FlowState::default();
    let mut measured = Vec::with_capacity(blocks.len());
    let mut dependencies = Vec::with_capacity(blocks.len());
    for (index, block) in blocks.iter_mut().enumerate() {
        let (extent, mut reads) = FontChainDependencies::capture(|| {
            flow.measure(
                index,
                block,
                widths,
                default_width,
                config,
                paragraph_zones,
                zones_by_anchor,
                section_break_marks,
            )
        });
        let unknown = FontChainDependencies::unknown();
        let zones = zone_dependencies.get(index).unwrap_or(&unknown);
        reads.extend(zones);
        zones.record();
        measured.push(extent?);
        dependencies.push(reads);
    }
    Ok((measured, dependencies))
}

/// Where the float flow stands between two blocks.
#[derive(Default)]
struct FlowState {
    cumulative_y: f64,
    active_zones: Vec<FloatingZone>,
}

impl FlowState {
    #[allow(clippy::too_many_arguments)]
    fn measure(
        &mut self,
        index: usize,
        block: &mut LayoutBlock,
        widths: &[f64],
        default_width: f64,
        config: &MeasurementConfig,
        paragraph_zones: &ParagraphZones,
        zones_by_anchor: &AnchorZones,
        section_break_marks: &[bool],
    ) -> Result<BlockExtent, String> {
        let Self {
            cumulative_y,
            active_zones,
        } = self;
        if resets_float_flow(block) {
            active_zones.clear();
            *cumulative_y = 0.0;
        }
        if let Some(zones) = paragraph_zones.get(&index) {
            // A paragraph-anchored band hangs off its own anchor, which sits at
            // `cumulative_y` in the frame the earlier bands were measured in.
            if active_zones.len() + zones.len() <= MAX_ACTIVE_ZONES {
                active_zones.extend(zones.iter().map(|zone| FloatingZone {
                    top_y: zone.top_y + *cumulative_y,
                    bottom_y: zone.bottom_y + *cumulative_y,
                    ..zone.clone()
                }));
            } else {
                *cumulative_y = 0.0;
                active_zones.clone_from(zones);
            }
        }
        if let Some(zones) = zones_by_anchor.get(&index) {
            // Anchors the flow has not advanced past share one origin.
            if *cumulative_y == 0.0 && active_zones.len() + zones.len() <= MAX_ACTIVE_ZONES {
                active_zones.extend(zones.iter().cloned());
            } else {
                *cumulative_y = 0.0;
                active_zones.clone_from(zones);
            }
        }
        let width = widths.get(index).copied().unwrap_or(default_width);
        // A bare paragraph mark carrying section properties is the section
        // break itself and prints no line, unless it is everything its section
        // holds: Word lays such a section out one line tall.
        let extent = if section_break_marks.get(index).copied().unwrap_or(false) {
            BlockExtent::Paragraph(ParagraphExtent {
                lines: Vec::new(),
                total_height: 0.0,
            })
        } else {
            measure_block_with_context(
                block,
                width,
                config,
                (!active_zones.is_empty()).then_some(active_zones.as_slice()),
                *cumulative_y,
            )?
        };
        if !matches!(block, LayoutBlock::Table(table) if table.floating.is_some())
            && !matches!(block, LayoutBlock::Shape(shape) if anchored_shape(shape))
        {
            *cumulative_y += extent_height(&extent);
        }
        Ok(extent)
    }
}

/// [`measure_blocks_with_floats`] a few blocks at a time. The floating zones
/// are found once up front; [`Self::measure_until`] then measures the blocks in
/// order and can stop after any of them, and the extents equal one call's.
pub struct FloatFlow {
    default_width: f64,
    paragraph_zones: ParagraphZones,
    zones_by_anchor: AnchorZones,
    marks: Vec<bool>,
    state: FlowState,
    measured: Vec<BlockExtent>,
    font_dependencies: Vec<FontChainDependencies>,
    zone_dependencies: Vec<FontChainDependencies>,
}

impl FloatFlow {
    pub fn new(
        blocks: &[LayoutBlock],
        widths: &[f64],
        config: &MeasurementConfig,
        page_geometry: Option<&FloatPageGeometry>,
    ) -> Result<Self, String> {
        Self::with_table_wrap_frames(blocks, widths, &[], config, page_geometry)
    }

    pub fn with_table_wrap_frames(
        blocks: &[LayoutBlock],
        widths: &[f64],
        table_wrap_frames: &[bool],
        config: &MeasurementConfig,
        page_geometry: Option<&FloatPageGeometry>,
    ) -> Result<Self, String> {
        let default_width = widths.first().copied().unwrap_or(0.0);
        let (extracted, zone_dependencies) = extract_floating_zones_recorded(
            blocks,
            default_width,
            widths,
            table_wrap_frames,
            config,
            page_geometry,
            &BTreeMap::new(),
            true,
        )?;
        let (paragraph_zones, zones_by_anchor) = group_floating_zones(extracted);
        Ok(Self {
            default_width,
            paragraph_zones,
            zones_by_anchor,
            marks: section_break_marks(blocks),
            state: FlowState::default(),
            measured: Vec::with_capacity(blocks.len()),
            font_dependencies: Vec::with_capacity(blocks.len()),
            zone_dependencies,
        })
    }

    /// Blocks measured so far.
    pub fn measured(&self) -> usize {
        self.measured.len()
    }

    /// The extents of the blocks measured so far.
    pub fn extents(&self) -> &[BlockExtent] {
        &self.measured
    }

    pub fn font_dependencies(&self) -> &[FontChainDependencies] {
        &self.font_dependencies
    }

    /// Measures the next blocks up to `end`, exclusive. `blocks` and `widths`
    /// are the ones the flow was created for.
    pub fn measure_until(
        &mut self,
        blocks: &mut [LayoutBlock],
        widths: &[f64],
        config: &MeasurementConfig,
        end: usize,
    ) -> Result<(), String> {
        for index in self.measured.len()..end.min(blocks.len()) {
            let (extent, mut dependencies) = FontChainDependencies::capture(|| {
                self.state.measure(
                    index,
                    &mut blocks[index],
                    widths,
                    self.default_width,
                    config,
                    &self.paragraph_zones,
                    &self.zones_by_anchor,
                    &self.marks,
                )
            });
            dependencies.extend(&self.zone_dependencies[index]);
            dependencies.record();
            self.measured.push(extent?);
            self.font_dependencies.push(dependencies);
        }
        Ok(())
    }

    pub fn into_extents(self) -> Vec<BlockExtent> {
        self.measured
    }
}

pub fn measure_block(
    block: &mut LayoutBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<BlockExtent, String> {
    match block {
        LayoutBlock::Paragraph(paragraph) => {
            measure_paragraph(paragraph, content_width, config).map(BlockExtent::Paragraph)
        }
        LayoutBlock::Table(table) => {
            measure_table(table, content_width, config).map(BlockExtent::Table)
        }
        LayoutBlock::Image(image) => Ok(BlockExtent::Image(ImageExtent {
            width: rotation_bound(&image.rotation_bounds, "width").unwrap_or(image.width),
            height: rotation_bound(&image.rotation_bounds, "height").unwrap_or(image.height),
        })),
        LayoutBlock::Shape(shape) => measure_shape(shape, config).map(BlockExtent::Shape),
        LayoutBlock::Chart(chart) => Ok(BlockExtent::Chart(ChartExtent {
            width: chart.width,
            height: chart.height,
        })),
        LayoutBlock::TextBox(text_box) => {
            let margins = text_box.margins.as_ref();
            let left = margins.map_or(7.0, |value| value.left);
            let right = margins.map_or(7.0, |value| value.right);
            let top = margins.map_or(4.0, |value| value.top);
            let bottom = margins.map_or(4.0, |value| value.bottom);
            let inner_width = (text_box.width - left - right).max(1.0);
            let inner_measures = text_box
                .content
                .iter()
                .map(|paragraph| measure_paragraph(paragraph, inner_width, config))
                .collect::<Result<Vec<_>, _>>()?;
            let content_height = inner_measures
                .iter()
                .map(|measure| measure.total_height)
                .sum::<f64>();
            Ok(BlockExtent::TextBox(TextBoxExtent {
                width: text_box.width,
                height: text_box.height.unwrap_or(content_height + top + bottom),
                inner_measures,
            }))
        }
        LayoutBlock::SectionBreak(_) => Ok(BlockExtent::SectionBreak),
        LayoutBlock::PageBreak(_) => Ok(BlockExtent::PageBreak),
        LayoutBlock::ColumnBreak(_) => Ok(BlockExtent::ColumnBreak),
        LayoutBlock::Unsupported => Ok(BlockExtent::Unsupported),
    }
}

fn measure_block_with_context(
    block: &mut LayoutBlock,
    content_width: f64,
    config: &MeasurementConfig,
    floating_zones: Option<&[FloatingZone]>,
    cumulative_y: f64,
) -> Result<BlockExtent, String> {
    match block {
        LayoutBlock::Paragraph(paragraph) => measure_paragraph_with_context(
            paragraph,
            content_width,
            config,
            floating_zones,
            cumulative_y,
        )
        .map(BlockExtent::Paragraph),
        _ => measure_block(block, content_width, config),
    }
}

fn rotation_bound(bounds: &Option<Value>, field: &str) -> Option<f64> {
    bounds.as_ref()?.get(field)?.as_f64()
}

pub(crate) fn measure_paragraph(
    paragraph: &ParagraphBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<ParagraphExtent, String> {
    measure_paragraph_with_context(paragraph, content_width, config, None, 0.0)
}

fn measure_paragraph_with_context(
    paragraph: &ParagraphBlock,
    content_width: f64,
    config: &MeasurementConfig,
    floating_zones: Option<&[FloatingZone]>,
    cumulative_y: f64,
) -> Result<ParagraphExtent, String> {
    let lookup = extent_cache_lookup(
        paragraph,
        content_width,
        config,
        floating_zones,
        cumulative_y,
    );
    if let ExtentLookup::Hit(extent) = lookup {
        return Ok(extent);
    }
    #[cfg(test)]
    EXTENT_MEASURE_CALLS.with(|calls| calls.set(calls.get() + 1));
    let (extent, dependencies) = FontChainDependencies::capture(|| {
        if !content_width.is_finite() || content_width <= 0.0 {
            synthetic_paragraph_extent(paragraph, content_width)
        } else {
            crate::typed_measure::measure_paragraph(
                paragraph,
                content_width,
                config,
                floating_zones,
                cumulative_y,
            )
            .unwrap_or_else(|| synthetic_paragraph_extent(paragraph, content_width))
        }
    });
    let mut extent = extent;
    measure_horizontal_rules(paragraph, &mut extent);
    if let ExtentLookup::Miss(Some(key)) = lookup {
        let weight = extent_weight(&extent) + dependencies.retained_bytes();
        EXTENT_CACHE.with(|cache| {
            cache
                .borrow_mut()
                .insert_hot(key, extent.clone(), dependencies, weight)
        });
    }
    Ok(extent)
}

const MAX_EXTENT_CACHE_ENTRIES: usize = 4_096;
const MAX_EXTENT_CACHE_KEY_BYTES: usize = 8 * 1024 * 1024;
/// Estimated retained bytes of cached extents per generation.
const MAX_EXTENT_CACHE_VALUE_BYTES: usize = 32 * 1024 * 1024;

fn extent_weight(extent: &ParagraphExtent) -> usize {
    use std::mem::size_of;
    size_of::<ParagraphExtent>()
        + extent.lines.capacity() * size_of::<TypesetRow>()
        + extent
            .lines
            .iter()
            .map(|row| {
                row.segments
                    .as_ref()
                    .map_or(0, |v| v.capacity() * size_of::<TypesetRowSegment>())
                    + row
                        .run_advances
                        .as_ref()
                        .map_or(0, |v| v.capacity() * size_of::<TypesetRunAdvance>())
                    + row
                        .cluster_advances
                        .as_ref()
                        .map_or(0, |v| v.capacity() * size_of::<TypesetClusterAdvance>())
                    + row
                        .bidi_slices
                        .as_ref()
                        .map_or(0, |v| v.capacity() * size_of::<TypesetBidiSlice>())
            })
            .sum::<usize>()
}
/// Serialized paragraph inputs past this size are measured uncached.
const MAX_EXTENT_KEY_BYTES: usize = 256 * 1024;

#[derive(Default)]
struct ExtentCacheGeneration {
    entries: HashMap<
        Vec<u8>,
        (ParagraphExtent, FontChainDependencies, usize),
        foldhash::fast::RandomState,
    >,
    key_bytes: usize,
    value_bytes: usize,
}

impl ExtentCacheGeneration {
    fn would_overflow(&self, key: &[u8], weight: usize) -> bool {
        self.entries.len() >= MAX_EXTENT_CACHE_ENTRIES
            || self.key_bytes.saturating_add(key.len()) > MAX_EXTENT_CACHE_KEY_BYTES
            || self.value_bytes.saturating_add(weight) > MAX_EXTENT_CACHE_VALUE_BYTES
    }

    fn insert(
        &mut self,
        key: Vec<u8>,
        extent: ParagraphExtent,
        dependencies: FontChainDependencies,
        weight: usize,
    ) {
        self.remove(&key);
        self.key_bytes += key.len();
        self.value_bytes += weight;
        self.entries.insert(key, (extent, dependencies, weight));
    }

    fn remove(&mut self, key: &[u8]) -> Option<(ParagraphExtent, FontChainDependencies, usize)> {
        let (extent, dependencies, weight) = self.entries.remove(key)?;
        self.key_bytes = self.key_bytes.saturating_sub(key.len());
        self.value_bytes = self.value_bytes.saturating_sub(weight);
        Some((extent, dependencies, weight))
    }
}

/// Measured paragraph extents reused across pagination passes; same
/// two-generation aging as `ooxml_text`'s shape cache.
#[derive(Default)]
struct ExtentCache {
    hot: ExtentCacheGeneration,
    cold: ExtentCacheGeneration,
}

impl ExtentCache {
    fn get(&mut self, key: &[u8], chains: FontChains<'_>) -> Option<ParagraphExtent> {
        if let Some((extent, dependencies, _)) = self.hot.entries.get(key) {
            if dependencies.matches(chains) {
                dependencies.record();
                return Some(extent.clone());
            }
            self.hot.remove(key);
        }
        let (extent, dependencies, weight) = self.cold.remove(key)?;
        if !dependencies.matches(chains) {
            return None;
        }
        dependencies.record();
        self.insert_hot(key.to_vec(), extent.clone(), dependencies, weight);
        Some(extent)
    }

    fn insert_hot(
        &mut self,
        key: Vec<u8>,
        extent: ParagraphExtent,
        dependencies: FontChainDependencies,
        weight: usize,
    ) {
        if self.hot.would_overflow(&key, weight) {
            self.cold = std::mem::take(&mut self.hot);
        }
        self.hot.insert(key, extent, dependencies, weight);
    }
}

thread_local! {
    static EXTENT_CACHE: RefCell<ExtentCache> = RefCell::new(ExtentCache::default());
    #[cfg(test)]
    static EXTENT_MEASURE_CALLS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    /// Key scratch reused per lookup so a hit allocates nothing.
    static EXTENT_KEY_BUF: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
}

pub(crate) fn clear_extent_cache() {
    let _ = EXTENT_CACHE.try_with(|cache| *cache.borrow_mut() = ExtentCache::default());
    let _ = EXTENT_KEY_BUF.try_with(|scratch| *scratch.borrow_mut() = Vec::new());
}

#[cfg(test)]
fn extent_cache_stats() -> (usize, usize) {
    let entries = EXTENT_CACHE.with(|cache| {
        let cache = cache.borrow();
        cache.hot.entries.len() + cache.cold.entries.len()
    });
    let scratch_capacity = EXTENT_KEY_BUF.with(|scratch| scratch.borrow().capacity());
    (entries, scratch_capacity)
}

enum ExtentLookup {
    Hit(ParagraphExtent),
    /// The built key to insert under, or `None` for uncached inputs.
    Miss(Option<Vec<u8>>),
}

/// Key over every input a measure reads; oversized inputs run uncached.
fn extent_cache_lookup(
    paragraph: &ParagraphBlock,
    content_width: f64,
    config: &MeasurementConfig,
    floating_zones: Option<&[FloatingZone]>,
    cumulative_y: f64,
) -> ExtentLookup {
    EXTENT_KEY_BUF.with(|scratch| {
        let key = &mut *scratch.borrow_mut();
        key.clear();
        if crate::extent_key::encode(key, paragraph).is_err() {
            return ExtentLookup::Miss(None);
        }
        if key.len() > MAX_EXTENT_KEY_BYTES {
            return ExtentLookup::Miss(None);
        }
        key.extend_from_slice(&content_width.to_bits().to_le_bytes());
        key.extend_from_slice(&cumulative_y.to_bits().to_le_bytes());
        match floating_zones {
            None => key.push(0),
            Some(zones) => {
                key.push(1);
                key.extend_from_slice(&(zones.len() as u64).to_le_bytes());
                for zone in zones {
                    for value in [
                        zone.left_margin,
                        zone.right_margin,
                        zone.top_y,
                        zone.bottom_y,
                    ] {
                        key.extend_from_slice(&value.to_bits().to_le_bytes());
                    }
                    key.push(zone.full_width_block as u8);
                    key.extend_from_slice(&(zone.segments.len() as u64).to_le_bytes());
                    for strip in &zone.segments {
                        key.extend_from_slice(&strip.left_offset.to_bits().to_le_bytes());
                        key.extend_from_slice(&strip.available_width.to_bits().to_le_bytes());
                    }
                }
            }
        }
        key.extend_from_slice(&config_fingerprint(config).to_le_bytes());
        let (store, availability) = crate::measure_font_cache_identity(&config.font_chains);
        key.extend_from_slice(&store.to_le_bytes());
        key.extend_from_slice(&availability.to_le_bytes());
        match EXTENT_CACHE.with(|cache| {
            cache
                .borrow_mut()
                .get(key, FontChains::BTree(&config.font_chains))
        }) {
            Some(extent) => ExtentLookup::Hit(extent),
            None => ExtentLookup::Miss(Some(key.clone())),
        }
    })
}

fn fnv1a(mut hash: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x00000100000001b3);
    }
    hash
}

fn config_fingerprint(config: &MeasurementConfig) -> u64 {
    let mut hash = fnv1a(0xcbf29ce484222325, &[config.authoritative_shaping as u8]);
    hash = fnv1a(hash, config.defaults.to_string().as_bytes());
    fnv1a(hash, config.compat.to_string().as_bytes())
}

fn measure_horizontal_rules(paragraph: &ParagraphBlock, extent: &mut ParagraphExtent) {
    let Some(attrs) = paragraph
        .attrs
        .as_ref()
        .filter(|attrs| !attrs.horizontal_rules.is_empty())
    else {
        return;
    };
    if attrs
        .spacing
        .as_ref()
        .and_then(|spacing| spacing.line_rule.as_deref())
        == Some("exact")
    {
        return;
    }
    for line in &mut extent.lines {
        let mut height = 0.0_f64;
        let mut standalone = true;
        let mut rule_count = 0;
        for (index, run) in paragraph
            .runs
            .iter()
            .enumerate()
            .take(line.tail_run + 1)
            .skip(line.head_run)
        {
            if index == line.tail_run && line.tail_char == 0 {
                continue;
            }
            if matches!(run, crate::types::Run::Text(text) if text.fmt.hidden == Some(true)) {
                continue;
            }
            if let Some(rule) = attrs
                .horizontal_rules
                .iter()
                .find(|rule| run.pm_start() == Some(rule.pm_start))
            {
                height = height.max(rule.height + 1.0);
                rule_count += 1;
            } else if !matches!(run, crate::types::Run::Text(text) if text.text.is_empty()) {
                standalone = false;
            }
        }
        if height <= 0.0 {
            continue;
        }
        let original_height = line.line_height;
        if standalone && rule_count == 1 {
            line.line_height = line.line_height.max(height);
            line.ascent = line.line_height;
            line.descent = 0.0;
        } else if height > line.ascent {
            let extra = height - line.ascent;
            line.ascent += extra;
            line.line_height += extra;
        }
        extent.total_height += line.line_height - original_height;
    }
}

const SYNTHETIC_ADVANCE_EM: f64 = 1.0;

fn valid_font_size(size: Option<f64>, fallback: f64) -> f64 {
    size.filter(|size| size.is_finite() && *size > 0.0)
        .unwrap_or(fallback)
}

fn synthetic_font_px(run: &crate::types::RunFormatting, default_font_size: f64) -> f64 {
    let size = valid_font_size(run.font_size, default_font_size);
    let script_scale = if run.superscript == Some(true) || run.subscript == Some(true) {
        0.75
    } else {
        1.0
    };
    size * 96.0 / 72.0 * script_scale
}

fn synthetic_scalar_count(text: &str, all_caps: Option<bool>) -> usize {
    if all_caps == Some(true) {
        text.chars().flat_map(char::to_uppercase).count()
    } else {
        text.chars().count()
    }
}

fn synthetic_text_width(
    text: &str,
    run: &crate::types::RunFormatting,
    default_font_size: f64,
) -> f64 {
    let scalars = synthetic_scalar_count(text, run.all_caps) as f64;
    let letter_spacing = run
        .letter_spacing
        .filter(|spacing| spacing.is_finite() && *spacing > 0.0 && *spacing <= 1_000.0)
        .unwrap_or(0.0);
    let horizontal_scale = run
        .horizontal_scale
        .filter(|scale| scale.is_finite() && *scale > 0.0 && *scale <= 600.0)
        .map(|scale| scale / 100.0)
        .unwrap_or(1.0);
    scalars
        * (synthetic_font_px(run, default_font_size) * SYNTHETIC_ADVANCE_EM + letter_spacing)
        * horizontal_scale
}

fn synthetic_inline_image_width(image: &crate::types::ImageRun) -> f64 {
    let floating = matches!(
        image.wrap_type.as_deref(),
        Some("square" | "tight" | "through" | "behind" | "inFront")
    ) || image.display_mode.as_deref() == Some("float");
    if floating {
        return 0.0;
    }
    rotation_bound(&image.rotation_bounds, "width")
        .unwrap_or(image.width)
        .max(0.0)
}

/// `w:spacing` mapped onto a line rule, in the same precedence order the
/// measured path uses. `None` is single spacing, which leaves the box alone.
fn synthetic_line_rule(spacing: Option<&ParagraphSpacing>) -> Option<LineSpacingRule> {
    let spacing = spacing?;
    match (
        spacing.line_rule.as_deref(),
        spacing.line,
        spacing.line_unit.as_deref(),
    ) {
        (Some("exact"), Some(line), _) => Some(LineSpacingRule::Exact {
            px: line.max(0.0) as f32,
        }),
        (Some("atLeast"), Some(line), _) => Some(LineSpacingRule::AtLeast {
            px: line.max(0.0) as f32,
        }),
        (_, Some(line), Some("multiplier")) => Some(LineSpacingRule::Auto {
            line_240ths: (line * 240.0).round().clamp(0.0, 24_000_000.0) as u32,
        }),
        (_, Some(line), Some("px")) => Some(LineSpacingRule::Exact {
            px: line.max(0.0) as f32,
        }),
        _ => None,
    }
}

fn synthetic_row(
    head_run: usize,
    tail_run: usize,
    tail_char: usize,
    width: f64,
    font_px: f64,
    rule: Option<&LineSpacingRule>,
) -> TypesetRow {
    let (ascent, descent, line_height) = match rule {
        // single spacing is the identity, so skip the f32 box round-trip
        None | Some(LineSpacingRule::Auto { line_240ths: 240 }) => {
            (font_px * 0.8, font_px * 0.2, font_px * 1.15)
        }
        Some(rule) => {
            let ruled = apply_spacing_rule(
                LineBox {
                    ascent: (font_px * 0.8) as f32,
                    descent: (font_px * 0.2) as f32,
                    leading: (font_px * 0.15) as f32,
                },
                rule,
            );
            (
                f64::from(ruled.ascent),
                f64::from(ruled.descent),
                f64::from(ruled.height()),
            )
        }
    };
    TypesetRow {
        head_run,
        head_char: 0,
        tail_run,
        tail_char,
        width,
        ascent,
        descent,
        line_height,
        synthetic_fallback: Some(true),
        marker_tab_offset: None,
        ..TypesetRow::default()
    }
}

fn synthetic_paragraph_extent(paragraph: &ParagraphBlock, content_width: f64) -> ParagraphExtent {
    let default_font_size = valid_font_size(
        paragraph
            .attrs
            .as_ref()
            .and_then(|attrs| attrs.default_font_size),
        11.0,
    );
    let default_font_px = default_font_size * 96.0 / 72.0;
    let spacing = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.spacing.as_ref());
    let rule = synthetic_line_rule(spacing);
    let measurable = content_width.is_finite() && content_width > 0.0;
    let slot = |width: f64| if measurable { width } else { 0.0 };

    let mut lines = Vec::new();
    let mut head_run = 0usize;
    let mut font_px = default_font_px;
    let mut line_width = 0.0f64;
    for (index, run) in paragraph.runs.iter().enumerate() {
        match run {
            Run::Text(text) => {
                font_px = font_px.max(synthetic_font_px(&text.fmt, default_font_size));
                if text.text != "\u{200b}" {
                    line_width += synthetic_text_width(&text.text, &text.fmt, default_font_size);
                }
            }
            Run::Tab(tab) => {
                font_px = font_px.max(synthetic_font_px(&tab.fmt, default_font_size));
                line_width += tab.width.unwrap_or(48.0).max(0.0);
            }
            Run::Image(image) => line_width += synthetic_inline_image_width(image),
            Run::Field(field) => {
                font_px = font_px.max(synthetic_font_px(&field.fmt, default_font_size));
                let fallback = field
                    .fallback
                    .as_deref()
                    .filter(|text| !text.is_empty())
                    .unwrap_or("1");
                line_width += synthetic_text_width(fallback, &field.fmt, default_font_size);
            }
            // an authored break is exact, so it splits the fallback rows even
            // though their widths are guesses
            Run::LineBreak(_) => {
                lines.push(synthetic_row(
                    head_run,
                    index,
                    0,
                    slot(line_width),
                    font_px,
                    rule.as_ref(),
                ));
                head_run = index + 1;
                font_px = default_font_px;
                line_width = 0.0;
            }
            Run::Unsupported => {}
        }
    }
    let tail_run = paragraph.runs.len().saturating_sub(1).max(head_run);
    let tail_char = paragraph.runs.get(tail_run).map_or(0, |run| match run {
        Run::Text(text) => text.text.encode_utf16().count(),
        _ => 0,
    });
    lines.push(synthetic_row(
        head_run,
        tail_run,
        tail_char,
        slot(line_width),
        font_px,
        rule.as_ref(),
    ));

    ParagraphExtent {
        total_height: spacing.and_then(|value| value.before).unwrap_or(0.0)
            + lines.iter().map(|line| line.line_height).sum::<f64>()
            + spacing.and_then(|value| value.after).unwrap_or(0.0),
        lines,
    }
}

fn extract_floating_zones(
    blocks: &[LayoutBlock],
    content_width: f64,
    widths: &[f64],
    table_wrap_frames: &[bool],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    shape_offsets: &BTreeMap<usize, f64>,
) -> Result<Vec<AnchoredFloatingZone>, String> {
    extract_floating_zones_recorded(
        blocks,
        content_width,
        widths,
        table_wrap_frames,
        config,
        page_geometry,
        shape_offsets,
        false,
    )
    .map(|(zones, _)| zones)
}

#[allow(clippy::too_many_arguments)]
fn extract_floating_zones_recorded(
    blocks: &[LayoutBlock],
    content_width: f64,
    widths: &[f64],
    table_wrap_frames: &[bool],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    shape_offsets: &BTreeMap<usize, f64>,
    record_dependencies: bool,
) -> Result<(Vec<AnchoredFloatingZone>, Vec<FontChainDependencies>), String> {
    let mut zones = Vec::new();
    let mut dependencies = if record_dependencies {
        Vec::with_capacity(blocks.len())
    } else {
        Vec::new()
    };
    for (block_index, block) in blocks.iter().enumerate() {
        let mut extract = || -> Result<(), String> {
            match block {
                LayoutBlock::Paragraph(paragraph) => {
                    extract_image_zones(paragraph, block_index, content_width, &mut zones);
                }
                LayoutBlock::Table(table) => {
                    extract_table_zone(
                        table,
                        block_index,
                        content_width,
                        widths
                            .get(block_index)
                            .copied()
                            .filter(|_| table_wrap_frames.get(block_index) == Some(&true)),
                        config,
                        &mut zones,
                    )?;
                }
                LayoutBlock::TextBox(text_box) => extract_text_box_zone(
                    text_box,
                    block_index,
                    content_width,
                    page_geometry,
                    &mut zones,
                ),
                LayoutBlock::Shape(shape) => extract_shape_zone(
                    shape,
                    block_index,
                    content_width,
                    page_geometry,
                    shape_offsets.get(&block_index).copied(),
                    &mut zones,
                ),
                _ => {}
            }
            Ok(())
        };
        if record_dependencies {
            let (result, reads) = FontChainDependencies::capture(extract);
            result?;
            dependencies.push(reads);
        } else {
            extract()?;
        }
    }
    Ok((zones, dependencies))
}

/// Whether a line runs past a float rather than stopping at its wider side.
fn flows_past(strip_left: f64, strip_right: f64, band: f64) -> bool {
    strip_left >= MIN_WRAP_SEGMENT_WIDTH
        && strip_right >= MIN_WRAP_SEGMENT_WIDTH
        && band < strip_left.min(strip_right)
}

fn extract_shape_zone(
    shape: &ShapeBlock,
    block_index: usize,
    content_width: f64,
    geometry: Option<&FloatPageGeometry>,
    resolved_x: Option<f64>,
    zones: &mut Vec<AnchoredFloatingZone>,
) {
    let Some(position) = shape.position.as_ref() else {
        return;
    };
    if !matches!(
        shape.wrap_type.as_deref(),
        Some("square" | "tight" | "through" | "topAndBottom")
    ) || shape.width <= 0.0
        || shape.height <= 0.0
    {
        return;
    }
    let margin_left = geometry.map_or(0.0, |g| g.margin_left);
    let page_width = geometry.map_or(content_width, |g| g.page_width);
    let horizontal = position.horizontal.as_ref();
    let page_relative = horizontal.and_then(|axis| axis.relative_to.as_deref()) == Some("page");
    let base_x = if page_relative { -margin_left } else { 0.0 };
    let frame_width = if page_relative {
        page_width
    } else {
        content_width
    };
    let x = resolved_x.unwrap_or_else(|| {
        base_x
            + match horizontal.and_then(|axis| axis.align.as_deref()) {
                Some("right" | "outside") => frame_width - shape.width,
                Some("center") => (frame_width - shape.width) / 2.0,
                Some("left" | "inside") => 0.0,
                _ => horizontal.and_then(|axis| axis.pos_offset).unwrap_or(0.0),
            }
    });
    let vertical = position.vertical.as_ref();
    let margin_top = geometry.map_or(0.0, |g| g.margin_top);
    let content_height = geometry.map_or(0.0, |g| g.content_height);
    let (base_y, frame_height) = match vertical.and_then(|axis| axis.relative_to.as_deref()) {
        Some("page") => (
            -margin_top,
            geometry.map_or(content_height, |g| g.page_height),
        ),
        Some("paragraph" | "line") => (0.0, 0.0),
        _ => (0.0, content_height),
    };
    let y = base_y
        + match vertical.and_then(|axis| axis.align.as_deref()) {
            Some("bottom") => frame_height - shape.height,
            Some("center") => (frame_height - shape.height) / 2.0,
            Some("top") => 0.0,
            _ => vertical.and_then(|axis| axis.pos_offset).unwrap_or(0.0),
        };
    let distances = shape.wrap_distances.as_ref();
    let left = x - distances.map_or(0.0, |d| d.left);
    let right = x + shape.width + distances.map_or(0.0, |d| d.right);
    let top_y = y - distances.map_or(0.0, |d| d.top);
    let bottom_y = y + shape.height + distances.map_or(0.0, |d| d.bottom);
    if right <= 0.0 || left >= content_width || bottom_y <= 0.0 {
        return;
    }
    let full_width_block = shape.wrap_type.as_deref() == Some("topAndBottom")
        || (left <= 0.0 && right >= content_width);
    let mut segments = Vec::new();
    let (left_margin, right_margin) = if full_width_block {
        (0.0, 0.0)
    } else {
        let strip_left = left.max(0.0);
        let strip_right = (content_width - right).max(0.0);
        let text_on_left = match shape.wrap_text.as_deref() {
            Some("left") => true,
            Some("right") => false,
            // `bothSides` is the schema default and the only value that puts
            // text past an interior float; `largest` keeps one side.
            None | Some("bothSides")
                if flows_past(strip_left, strip_right, (right - left).max(0.0)) =>
            {
                segments = vec![
                    FloatStrip {
                        left_offset: 0.0,
                        available_width: strip_left,
                    },
                    FloatStrip {
                        left_offset: right.max(0.0),
                        available_width: strip_right,
                    },
                ];
                false
            }
            _ => strip_left > strip_right,
        };
        if !segments.is_empty() {
            (0.0, 0.0)
        } else if text_on_left {
            (0.0, (content_width - left).max(0.0))
        } else {
            (right.max(0.0), 0.0)
        }
    };
    zones.push(AnchoredFloatingZone {
        zone: FloatingZone {
            left_margin,
            right_margin,
            top_y,
            bottom_y,
            segments,
            full_width_block,
        },
        anchor_block_index: block_index,
        margin_relative: is_margin_relative(Some(position)),
    });
}

fn extract_image_zones(
    paragraph: &ParagraphBlock,
    block_index: usize,
    content_width: f64,
    zones: &mut Vec<AnchoredFloatingZone>,
) {
    for run in &paragraph.runs {
        let Run::Image(image) = run else {
            continue;
        };
        let wraps = matches!(
            image.wrap_type.as_deref(),
            Some("square" | "tight" | "through")
        ) || (image.display_mode.as_deref() == Some("float")
            && image.css_float.as_deref() != Some("none"));
        if !wraps || image.wrap_type.as_deref() == Some("topAndBottom") {
            continue;
        }
        let vertical = image
            .position
            .as_ref()
            .and_then(|value| value.vertical.as_ref());
        let top_y = if vertical.is_some_and(|value| {
            value.align.as_deref() == Some("top") && value.relative_to.as_deref() == Some("margin")
        }) {
            0.0
        } else {
            vertical
                .and_then(|value| value.pos_offset)
                .map_or(0.0, emu_to_pixels)
        };
        let (left_margin, right_margin) = anchored_margins(
            image.position.as_ref(),
            image.css_float.as_deref(),
            image.width,
            image.dist_left.unwrap_or(12.0),
            image.dist_right.unwrap_or(12.0),
            content_width,
        );
        if left_margin <= 0.0 && right_margin <= 0.0 {
            continue;
        }
        zones.push(AnchoredFloatingZone {
            zone: FloatingZone {
                left_margin,
                right_margin,
                top_y: top_y - image.dist_top.unwrap_or(0.0),
                bottom_y: top_y + image.height + image.dist_bottom.unwrap_or(0.0),
                segments: Vec::new(),
                full_width_block: false,
            },
            anchor_block_index: block_index,
            margin_relative: is_margin_relative(image.position.as_ref()),
        });
    }
}

/// Per block, whether its float flow (up to the next [`resets_float_flow`]
/// block) holds a paragraph with a negative indent. Painting lets such a line
/// run past a wrap margin, so its tables keep main's wrap side: a frame flag
/// passed to [`measure_blocks_with_table_wrap_frames`] must be false there.
pub fn negative_indent_float_flows(blocks: &[&LayoutBlock]) -> Vec<bool> {
    let mut flows = vec![false; blocks.len()];
    let mut start = 0;
    for index in 1..=blocks.len() {
        if index == blocks.len() || resets_float_flow(blocks[index]) {
            let negative = blocks[start..index]
                .iter()
                .any(|block| has_negative_side_indent(block));
            flows[start..index].fill(negative);
            start = index;
        }
    }
    flows
}

fn has_negative_side_indent(block: &LayoutBlock) -> bool {
    let LayoutBlock::Paragraph(paragraph) = block else {
        return false;
    };
    let Some(indent) = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.indent.as_ref())
    else {
        return false;
    };
    let left = indent.left.unwrap_or(0.0);
    let first = left + indent.first_line.unwrap_or(0.0) - indent.hanging.unwrap_or(0.0);
    left < 0.0 || first < 0.0 || indent.right.unwrap_or(0.0) < 0.0
}

fn extract_table_zone(
    table: &TableBlock,
    block_index: usize,
    content_width: f64,
    column_width: Option<f64>,
    config: &MeasurementConfig,
    zones: &mut Vec<AnchoredFloatingZone>,
) -> Result<(), String> {
    if table.floating.is_none() {
        return Ok(());
    }
    let mut measured_table = table.clone();
    let measure = measure_table(&mut measured_table, content_width, config)?;
    if let Some(mut zone) = table_floating_zone(table, &measure, content_width, column_width) {
        (zone.left_margin, zone.right_margin) =
            clamp_margins(zone.left_margin, zone.right_margin, content_width);
        zones.push(AnchoredFloatingZone {
            zone,
            anchor_block_index: block_index,
            margin_relative: false,
        });
    }
    Ok(())
}

fn table_floating_zone(
    table: &TableBlock,
    measure: &TableExtent,
    content_width: f64,
    column_width: Option<f64>,
) -> Option<FloatingZone> {
    let floating = table.floating.as_ref()?;
    let x = if let Some(value) = floating.tblp_x {
        value
    } else {
        match floating.tblp_x_spec.as_deref() {
            Some("right" | "outside") => content_width - measure.total_width,
            Some("center") => (content_width - measure.total_width) / 2.0,
            Some("left" | "inside") => 0.0,
            _ if table.justification.as_deref() == Some("center") => {
                (content_width - measure.total_width) / 2.0
            }
            _ if table.justification.as_deref() == Some("right") => {
                content_width - measure.total_width
            }
            _ => 0.0,
        }
    };
    Some(table_floating_zone_at_x(
        floating,
        measure,
        content_width,
        x,
        column_width,
    ))
}

fn table_floating_zone_at_x(
    floating: &FloatingTablePosition,
    measure: &TableExtent,
    content_width: f64,
    x: f64,
    column_width: Option<f64>,
) -> FloatingZone {
    let margin_right_of_table = x + measure.total_width + floating.right_from_text.unwrap_or(12.0);
    let margin_left_of_table = content_width - x + floating.left_from_text.unwrap_or(12.0);
    let main_text_on_right = x < content_width / 2.0;
    let text_on_right = column_width
        .filter(|width| width.is_finite() && *width > 0.0 && *width == content_width)
        .filter(|_| {
            matches!(
                floating.horz_anchor.as_deref(),
                None | Some("text" | "margin")
            ) && !matches!(floating.tblp_x_spec.as_deref(), Some("inside" | "outside"))
                && match floating.tblp_x {
                    Some(offset) => offset.is_finite(),
                    None => matches!(
                        floating.tblp_x_spec.as_deref(),
                        Some("left" | "right" | "center")
                    ),
                }
        })
        .filter(|width| measure.total_width > *width / 2.0)
        // Only where the side main picks leaves no room at all, so its text
        // runs full width under the table.
        .filter(|width| {
            let main_margin = if main_text_on_right {
                margin_right_of_table
            } else {
                margin_left_of_table
            };
            main_margin >= width.max(1.0)
        })
        .and_then(|width| {
            let (left_space, right_space) =
                table_wrap_gaps(floating, measure.total_width, width, x);
            (left_space >= MIN_WRAP_SEGMENT_WIDTH || right_space >= MIN_WRAP_SEGMENT_WIDTH)
                .then_some(right_space >= left_space)
        })
        .unwrap_or(main_text_on_right);
    let (left_margin, right_margin) = if text_on_right {
        (margin_right_of_table, 0.0)
    } else {
        (0.0, margin_left_of_table)
    };
    let top_y = floating.tblp_y.unwrap_or(0.0);
    FloatingZone {
        left_margin,
        right_margin,
        top_y: top_y - floating.top_from_text.unwrap_or(0.0),
        bottom_y: top_y + measure.total_height + floating.bottom_from_text.unwrap_or(0.0),
        segments: Vec::new(),
        full_width_block: false,
    }
}

fn extract_text_box_zone(
    text_box: &TextBoxBlock,
    block_index: usize,
    content_width: f64,
    page_geometry: Option<&FloatPageGeometry>,
    zones: &mut Vec<AnchoredFloatingZone>,
) {
    if text_box.display_mode.as_deref() != Some("float")
        && !matches!(
            text_box.wrap_type.as_deref(),
            Some("square" | "tight" | "through" | "behind" | "inFront" | "topAndBottom")
        )
    {
        return;
    }
    if matches!(text_box.wrap_type.as_deref(), Some("behind" | "inFront")) {
        return;
    }
    let height = text_box.height.unwrap_or(0.0);
    if text_box.width <= 0.0 || height <= 0.0 {
        return;
    }
    let margin_relative = is_margin_relative(text_box.position.as_ref());
    if text_box.wrap_type.as_deref() == Some("topAndBottom") {
        let raw_top = anchored_vertical_top(text_box.position.as_ref(), height, page_geometry);
        let bottom_y = raw_top + height + text_box.dist_bottom.unwrap_or(0.0);
        if bottom_y <= 0.0 {
            return;
        }
        zones.push(AnchoredFloatingZone {
            zone: FloatingZone {
                left_margin: 0.0,
                right_margin: 0.0,
                top_y: (raw_top - text_box.dist_top.unwrap_or(0.0)).max(0.0),
                bottom_y,
                segments: Vec::new(),
                full_width_block: true,
            },
            anchor_block_index: block_index,
            margin_relative,
        });
        return;
    }
    let top_y = text_box
        .position
        .as_ref()
        .and_then(|value| value.vertical.as_ref())
        .and_then(|value| value.pos_offset)
        .map_or(0.0, emu_to_pixels);
    let (left_margin, right_margin) = anchored_margins(
        text_box.position.as_ref(),
        text_box.css_float.as_deref(),
        text_box.width,
        text_box.dist_left.unwrap_or(12.0),
        text_box.dist_right.unwrap_or(12.0),
        content_width,
    );
    if left_margin <= 0.0 && right_margin <= 0.0 {
        return;
    }
    zones.push(AnchoredFloatingZone {
        zone: FloatingZone {
            left_margin,
            right_margin,
            top_y: top_y - text_box.dist_top.unwrap_or(0.0),
            bottom_y: top_y + height + text_box.dist_bottom.unwrap_or(0.0),
            segments: Vec::new(),
            full_width_block: false,
        },
        anchor_block_index: block_index,
        margin_relative,
    });
}

fn anchored_margins(
    position: Option<&ImageRunPosition>,
    css_float: Option<&str>,
    width: f64,
    dist_left: f64,
    dist_right: f64,
    content_width: f64,
) -> (f64, f64) {
    let horizontal = position.and_then(|value| value.horizontal.as_ref());
    let (left, right) = if horizontal.and_then(|value| value.align.as_deref()) == Some("left") {
        (width + dist_right, 0.0)
    } else if horizontal.and_then(|value| value.align.as_deref()) == Some("right") {
        (0.0, width + dist_left)
    } else if let Some(offset) = horizontal.and_then(|value| value.pos_offset) {
        let x = emu_to_pixels(offset);
        if x < content_width / 2.0 {
            (x + width + dist_right, 0.0)
        } else {
            (0.0, content_width - x + dist_left)
        }
    } else if css_float == Some("left") {
        (width + dist_right, 0.0)
    } else if css_float == Some("right") {
        (0.0, width + dist_left)
    } else {
        (0.0, 0.0)
    };
    clamp_margins(left, right, content_width)
}

fn clamp_margins(left: f64, right: f64, content_width: f64) -> (f64, f64) {
    let width = content_width.max(1.0);
    let left = left.max(0.0);
    let right = right.max(0.0);
    if left >= width || right >= width || left + right >= width {
        (0.0, 0.0)
    } else {
        (left, right)
    }
}

fn is_margin_relative(position: Option<&ImageRunPosition>) -> bool {
    matches!(
        position
            .and_then(|value| value.vertical.as_ref())
            .and_then(|value| value.relative_to.as_deref()),
        Some("margin" | "page")
    )
}

fn anchored_vertical_top(
    position: Option<&ImageRunPosition>,
    height: f64,
    geometry: Option<&FloatPageGeometry>,
) -> f64 {
    let Some(vertical) = position.and_then(|value| value.vertical.as_ref()) else {
        return 0.0;
    };
    let page_height = geometry.map_or(0.0, |value| value.page_height);
    let margin_top = geometry.map_or(0.0, |value| value.margin_top);
    let content_height = geometry.map_or(0.0, |value| value.content_height);
    let (base, size) = match vertical.relative_to.as_deref() {
        Some("paragraph" | "line") => (0.0, 0.0),
        Some("page") => (-margin_top, page_height),
        Some("topMargin") => (-margin_top, margin_top),
        Some("bottomMargin") => (content_height, margin_top),
        _ => (0.0, content_height),
    };
    match vertical.align.as_deref() {
        Some("top") => base,
        Some("center") if size != 0.0 => base + (size - height) / 2.0,
        Some("bottom") if size != 0.0 => base + size - height,
        _ if vertical.pos_offset.is_some() => {
            base + emu_to_pixels(vertical.pos_offset.unwrap_or(0.0))
        }
        _ if matches!(vertical.relative_to.as_deref(), Some("paragraph" | "line")) => 0.0,
        _ => base,
    }
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 9_525.0
}

/// One inset of the shape's text body in px, as the display list reads it, so
/// the wrap width measured here is the width the text is later emitted into.
fn text_body_inset(shape: &ShapeBlock, side: &str) -> f64 {
    shape
        .text_body_properties
        .as_ref()
        .and_then(|properties| properties.get("margins"))
        .and_then(|margins| margins.get(side))
        .and_then(Value::as_f64)
        .unwrap_or(0.0)
}

fn measure_shape(
    shape: &mut ShapeBlock,
    config: &MeasurementConfig,
) -> Result<ShapeExtent, String> {
    let inner_width =
        (shape.width - text_body_inset(shape, "left") - text_body_inset(shape, "right")).max(1.0);
    let inner_measures = shape
        .inner_text
        .as_ref()
        .map(|paragraphs| {
            paragraphs
                .iter()
                .map(|paragraph| measure_paragraph(paragraph, inner_width, config))
                .collect::<Result<Vec<_>, _>>()
        })
        .transpose()?
        .unwrap_or_default();
    shape.inner_measures = Some(inner_measures.clone());
    for child in &mut shape.children {
        measure_shape(child, config)?;
    }
    Ok(ShapeExtent {
        width: shape.width,
        height: shape.height,
        inner_measures: Some(inner_measures),
    })
}

fn measure_cell_blocks_with_table_floats(
    blocks: &mut [LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<Vec<BlockExtent>, String> {
    let mut measured = Vec::with_capacity(blocks.len());
    let mut zones = Vec::new();
    let mut y = 0.0_f64;
    let mut previous_after = 0.0_f64;
    for block in blocks {
        let spacing = match &*block {
            LayoutBlock::Paragraph(paragraph) => paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.spacing.as_ref()),
            _ => None,
        };
        let before = spacing.and_then(|spacing| spacing.before).unwrap_or(0.0);
        let after = spacing.and_then(|spacing| spacing.after).unwrap_or(0.0);
        // `y` tracks the paragraph's top, space-before included, which is the
        // origin the measurer probes float zones from.
        y += (previous_after - before).max(0.0);
        let extent = match &*block {
            LayoutBlock::Paragraph(paragraph) => {
                BlockExtent::Paragraph(measure_paragraph_with_context(
                    paragraph,
                    content_width,
                    config,
                    (!zones.is_empty()).then_some(zones.as_slice()),
                    y,
                )?)
            }
            _ => measure_block(block, content_width, config)?,
        };
        if let (LayoutBlock::Table(table), BlockExtent::Table(measure)) = (&*block, &extent)
            && let Some(floating) = table.floating.as_ref()
            && nested_table_float_offset(table.floating.as_ref()).is_some()
        {
            let x = nested_table_horizontal_offset(
                Some(floating),
                table.justification.as_deref(),
                table.indent,
                table.compatibility_mode,
                table.cell_margin_left,
                measure.total_width,
                content_width,
            );
            let mut zone = table_floating_zone_at_x(floating, measure, content_width, x, None);
            zone.top_y += y;
            zone.bottom_y += y;
            zones.push(zone);
        } else {
            y += extent_height(&extent) - after;
        }
        previous_after = after;
        measured.push(extent);
    }
    Ok(measured)
}

/// The width a paragraph wants when nothing forces it to wrap: its widest
/// typeset line plus the indents that sit beside it.
fn paragraph_content_width(
    paragraph: &crate::types::ParagraphBlock,
    budget: f64,
    config: &MeasurementConfig,
) -> Result<f64, String> {
    let indent = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.indent.as_ref());
    let edge = |value: Option<f64>| value.unwrap_or(0.0).max(0.0);
    let first_line = indent.map_or(0.0, |indent| edge(indent.first_line));
    let extent = measure_paragraph(paragraph, budget, config)?;
    let widest = extent
        .lines
        .iter()
        .enumerate()
        .map(|(index, line)| line.width + if index == 0 { first_line } else { 0.0 })
        .fold(0.0_f64, f64::max);
    Ok(widest + indent.map_or(0.0, |indent| edge(indent.left) + edge(indent.right)))
}

/// Per-column widest unwrapped content, for `columns` only; every other entry
/// stays zero so [`grow_content_sized_columns`] leaves it alone.
fn column_content_maximums(
    table: &TableBlock,
    columns: &[usize],
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<Vec<f64>, String> {
    let mut maximums = vec![0.0_f64; count_table_columns(table)];
    for entry in resolve_cell_grid(table) {
        if entry.col_span != 1 || !columns.contains(&entry.column_index) {
            continue;
        }
        let Some(cell) = table
            .rows
            .get(entry.row_index)
            .and_then(|row| row.cells.get(entry.cell_index))
        else {
            continue;
        };
        let left = cell
            .padding
            .as_ref()
            .map_or(DEFAULT_CELL_PADDING_X, |padding| padding.left);
        let right = cell
            .padding
            .as_ref()
            .map_or(DEFAULT_CELL_PADDING_X, |padding| padding.right);
        let budget = (content_width - left - right).max(1.0);
        let mut widest = 0.0_f64;
        for block in &cell.blocks {
            if let LayoutBlock::Paragraph(paragraph) = block {
                widest = widest.max(paragraph_content_width(paragraph, budget, config)?);
            }
        }
        if widest > 0.0
            && let Some(slot) = maximums.get_mut(entry.column_index)
        {
            *slot = slot.max(widest + left + right);
        }
    }
    Ok(maximums)
}

/// Measures every cell's blocks at its columns' width, a rotated cell at its
/// row's height; heights stay zero for the caller to settle.
fn measure_table_cells(
    table: &mut TableBlock,
    grid: &[ResolvedGridCell],
    column_widths: &[f64],
    content_width: f64,
    target_width: f64,
    config: &MeasurementConfig,
) -> Result<Vec<TableRowExtent>, String> {
    let mut rows = Vec::with_capacity(table.rows.len());

    for (row_index, row) in table.rows.iter_mut().enumerate() {
        let mut cells = Vec::with_capacity(row.cells.len());
        for (cell_index, cell) in row.cells.iter_mut().enumerate() {
            let resolved = grid
                .iter()
                .find(|entry| entry.row_index == row_index && entry.cell_index == cell_index);
            let column_index = resolved.map_or(0, |entry| entry.column_index);
            let col_span = cell.col_span.unwrap_or(1.0).max(1.0) as usize;
            let mut cell_width = column_widths
                .iter()
                .skip(column_index)
                .take(col_span)
                .sum::<f64>();
            if cell_width == 0.0 {
                cell_width = cell
                    .width
                    .filter(|width| *width > 0.0)
                    .or_else(|| {
                        resolve_table_width_px(
                            cell.width_value,
                            cell.width_type.as_deref(),
                            target_width,
                        )
                    })
                    .unwrap_or(100.0);
            }
            let left = cell
                .padding
                .as_ref()
                .map_or(DEFAULT_CELL_PADDING_X, |padding| padding.left);
            let right = cell
                .padding
                .as_ref()
                .map_or(DEFAULT_CELL_PADDING_X, |padding| padding.right);
            let rotated = is_rotated(cell);
            let measure_width = if rotated {
                let padding = cell
                    .padding
                    .as_ref()
                    .map_or(0.0, |padding| padding.top + padding.bottom);
                row.height.unwrap_or(content_width) - padding
            } else {
                cell_width - left - right
            };
            let has_table_floats = cell.blocks.iter().any(|block| {
                matches!(block, LayoutBlock::Table(table)
                    if nested_table_float_offset(table.floating.as_ref()).is_some())
            });
            let measures = if has_table_floats {
                measure_cell_blocks_with_table_floats(
                    &mut cell.blocks,
                    measure_width.max(1.0),
                    config,
                )?
            } else {
                measure_blocks(&mut cell.blocks, measure_width.max(1.0), config)?
            };
            cells.push(TableCellExtent {
                blocks: measures,
                width: cell_width,
                height: 0.0,
                col_span: cell.col_span,
                row_span: cell.row_span,
            });
        }
        rows.push(TableRowExtent { cells, height: 0.0 });
    }
    Ok(rows)
}

/// Whether a line of some measured cell paragraph starts inside a word.
fn table_breaks_inside_a_word(table: &TableBlock, rows: &[TableRowExtent]) -> bool {
    table.rows.iter().zip(rows).any(|(row, measured)| {
        row.cells
            .iter()
            .zip(&measured.cells)
            .any(|(cell, extent)| cell_breaks_inside_a_word(cell, extent))
    })
}

fn cell_breaks_inside_a_word(cell: &crate::types::TableCell, extent: &TableCellExtent) -> bool {
    !is_rotated(cell)
        && cell
            .blocks
            .iter()
            .zip(&extent.blocks)
            .any(|pair| match pair {
                (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent)) => {
                    starts_a_line_inside_a_word(paragraph, extent)
                }
                _ => false,
            })
}

/// Whether a line of `extent` starts where the paragraph's text offers no
/// line break: inside a word too wide for the line, or at the seam of two
/// runs that split one word.
fn starts_a_line_inside_a_word(paragraph: &ParagraphBlock, extent: &ParagraphExtent) -> bool {
    let text = |index: usize| match paragraph.runs.get(index) {
        Some(Run::Text(run)) if !run.text.is_empty() => Some(run.text.as_str()),
        _ => None,
    };
    let mut cursor = (usize::MAX, 0usize, 0usize);
    let mut opportunities: (usize, Vec<usize>) = (usize::MAX, Vec::new());
    extent.lines.iter().skip(1).any(|line| {
        let Some(current) = text(line.head_run) else {
            return false;
        };
        if cursor.0 != line.head_run || cursor.2 > line.head_char {
            cursor = (line.head_run, 0, 0);
        }
        let (_, mut byte, mut units) = cursor;
        while units < line.head_char {
            let Some(character) = current[byte..].chars().next() else {
                break;
            };
            units += character.len_utf16();
            byte += character.len_utf8();
        }
        cursor = (line.head_run, byte, units);
        let Some(character) = current[byte..]
            .chars()
            .next()
            .filter(|_| units == line.head_char)
        else {
            return false;
        };
        let Some(before) = current[..byte].chars().next_back() else {
            let previous = paragraph.runs[..line.head_run]
                .iter()
                .rev()
                .find(|candidate| !matches!(candidate, Run::Text(run) if run.text.is_empty()));
            return matches!(previous, Some(Run::Text(previous)) if previous
                .text
                .chars()
                .next_back()
                .is_some_and(|last| !ooxml_text::break_allowed_between(last, character)));
        };
        if breaking_space(before) || breaking_space(character) {
            return false;
        }
        if opportunities.0 != line.head_run {
            opportunities = (
                line.head_run,
                ooxml_text::break_opportunities(current)
                    .iter()
                    .map(|opportunity| opportunity.byte_index)
                    .collect(),
            );
        }
        opportunities.1.binary_search(&byte).is_err()
    })
}

/// Whitespace a line may break at; no-break spaces keep their word together.
fn breaking_space(character: char) -> bool {
    character.is_whitespace() && !matches!(character, '\u{00A0}' | '\u{2007}' | '\u{202F}')
}

/// The narrowest width that holds the widest word of the cell at `entry`,
/// margins included: `Some(None)` when the cell holds anything but
/// paragraphs, `None` when the entry names no cell. A rotated cell's lines
/// wrap along its row, so it asks for no room: it holds its stacked lines up
/// to its current width.
fn cell_content_minimum(
    table: &TableBlock,
    entry: &ResolvedGridCell,
    rows: &[TableRowExtent],
    content_width: f64,
    config: &MeasurementConfig,
) -> Option<Option<f64>> {
    let cell = table
        .rows
        .get(entry.row_index)
        .and_then(|row| row.cells.get(entry.cell_index))?;
    let padding = cell_horizontal_padding(cell);
    if is_rotated(cell) {
        return Some(
            rows.get(entry.row_index)
                .and_then(|row| row.cells.get(entry.cell_index))
                .map(|measured| {
                    (measured.blocks.iter().map(extent_height).sum::<f64>() + padding)
                        .min(measured.width)
                }),
        );
    }
    let content = cell
        .blocks
        .iter()
        .try_fold(0.0_f64, |widest, block| match block {
            LayoutBlock::Paragraph(paragraph) => {
                crate::typed_measure::min_content_width(paragraph, content_width, config)
                    .map(|width| widest.max(width))
            }
            _ => None,
        });
    Some(content.map(|content| content + padding))
}

/// Whether a cell spanning several columns breaks a line inside a word in
/// `after` where it did not in `before`, or, rotated, no longer holds its
/// stacked lines.
fn spanning_cell_newly_breaks_a_word(
    table: &TableBlock,
    grid: &[ResolvedGridCell],
    before: (&[f64], &[TableRowExtent]),
    after: (&[f64], &[TableRowExtent]),
) -> bool {
    fn measured<'a>(
        rows: &'a [TableRowExtent],
        entry: &ResolvedGridCell,
    ) -> Option<&'a TableCellExtent> {
        rows.get(entry.row_index)
            .and_then(|row| row.cells.get(entry.cell_index))
    }
    grid.iter().filter(|entry| entry.col_span > 1).any(|entry| {
        let Some(cell) = table
            .rows
            .get(entry.row_index)
            .and_then(|row| row.cells.get(entry.cell_index))
        else {
            return false;
        };
        if is_rotated(cell) {
            let overflows = |(widths, rows): (&[f64], &[TableRowExtent])| {
                let span: f64 = widths
                    .iter()
                    .skip(entry.column_index)
                    .take(entry.col_span)
                    .sum();
                measured(rows, entry).is_some_and(|extent| {
                    extent.blocks.iter().map(extent_height).sum::<f64>()
                        + cell_horizontal_padding(cell)
                        > span
                })
            };
            return overflows(after) && !overflows(before);
        }
        cell.blocks.iter().enumerate().any(|(index, block)| {
            let LayoutBlock::Paragraph(paragraph) = block else {
                return false;
            };
            let breaks = |rows: &[TableRowExtent]| {
                matches!(
                    measured(rows, entry).and_then(|extent| extent.blocks.get(index)),
                    Some(BlockExtent::Paragraph(extent))
                        if starts_a_line_inside_a_word(paragraph, extent)
                )
            };
            breaks(after.1) && !breaks(before.1)
        })
    })
}

fn cell_horizontal_padding(cell: &crate::types::TableCell) -> f64 {
    cell.padding
        .as_ref()
        .map_or(2.0 * DEFAULT_CELL_PADDING_X, |padding| {
            padding.left + padding.right
        })
}

/// Per column, the narrowest width that holds the widest word of every cell
/// sitting in that column alone, margins included; a rotated cell holds its
/// stacked lines up to its current width. A column that holds anything but paragraphs, or that only
/// spanning cells cover, is pinned with NaN.
fn column_content_minimums(
    table: &TableBlock,
    grid: &[ResolvedGridCell],
    rows: &[TableRowExtent],
    content_width: f64,
    config: &MeasurementConfig,
) -> Vec<f64> {
    let count = count_table_columns(table);
    let mut minimums = vec![f64::NAN; count];
    let mut pinned = vec![false; count];
    for entry in grid {
        if entry.col_span != 1 || entry.column_index >= count {
            continue;
        }
        let column = entry.column_index;
        match cell_content_minimum(table, entry, rows, content_width, config) {
            Some(Some(width)) => {
                minimums[column] = if minimums[column].is_nan() {
                    width
                } else {
                    minimums[column].max(width)
                };
            }
            Some(None) => pinned[column] = true,
            None => {}
        }
    }
    for (minimum, pinned) in minimums.iter_mut().zip(pinned) {
        if pinned {
            *minimum = f64::NAN;
        }
    }
    minimums
}

fn measure_table(
    table: &mut TableBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<TableExtent, String> {
    measure_table_with_compat_shift(table, content_width, config, true)
}

fn measure_table_with_compat_shift(
    table: &mut TableBlock,
    content_width: f64,
    config: &MeasurementConfig,
    apply_compat_shift: bool,
) -> Result<TableExtent, String> {
    let percentage_basis = if apply_compat_shift {
        table_percentage_basis(table, content_width)
    } else {
        content_width
    };
    let explicit_width =
        resolve_table_width_px(table.width, table.width_type.as_deref(), percentage_basis);
    let target_width = explicit_width.unwrap_or(content_width);
    let mut column_widths =
        resolve_table_column_widths_with_percentage_basis(table, content_width, percentage_basis);
    let content_sized = content_sized_columns(table, content_width, &column_widths);
    if !content_sized.is_empty() {
        let maximums = column_content_maximums(table, &content_sized, content_width, config)?;
        grow_content_sized_columns(table, content_width, &maximums, &mut column_widths);
    }
    let grid = resolve_cell_grid(table);
    let mut rows = measure_table_cells(
        table,
        &grid,
        &column_widths,
        content_width,
        target_width,
        config,
    )?;
    if fits_columns_to_words(table) && table_breaks_inside_a_word(table, &rows) {
        let before = column_widths.clone();
        if widen_columns_to_minimums(
            &mut column_widths,
            &column_content_minimums(table, &grid, &rows, content_width, config),
        ) {
            let widened = measure_table_cells(
                table,
                &grid,
                &column_widths,
                content_width,
                target_width,
                config,
            )?;
            if spanning_cell_newly_breaks_a_word(
                table,
                &grid,
                (&before, &rows),
                (&column_widths, &widened),
            ) {
                column_widths = before;
            } else {
                rows = widened;
            }
        }
    }

    let mut exact = vec![false; rows.len()];
    for (row_index, measured_row) in rows.iter_mut().enumerate() {
        let source_row = &table.rows[row_index];
        let mut max_height = 0.0_f64;
        let mut max_padding_height = 0.0_f64;
        let mut max_border_height = 0.0_f64;
        for (cell_index, measured_cell) in measured_row.cells.iter_mut().enumerate() {
            let source_cell = &source_row.cells[cell_index];
            let has_table_floats = source_cell.blocks.iter().any(|block| {
                matches!(block, LayoutBlock::Table(table)
                    if nested_table_float_offset(table.floating.as_ref()).is_some())
            });
            let mut content_height = 0.0_f64;
            let mut previous_after = 0.0_f64;
            let mut float_bottom = 0.0_f64;
            for (block, measure) in source_cell.blocks.iter().zip(&measured_cell.blocks) {
                if let (LayoutBlock::Table(table), BlockExtent::Table(extent)) = (block, measure)
                    && let Some(offset) = nested_table_float_offset(table.floating.as_ref())
                {
                    content_height += previous_after;
                    float_bottom = float_bottom.max(content_height + offset + extent.total_height);
                    previous_after = 0.0;
                    continue;
                }
                if let LayoutBlock::Shape(shape) = block
                    && crate::cell_layout::cell_overlay_drawing(
                        shape.position.is_some(),
                        shape.wrap_type.as_deref(),
                    )
                {
                    continue;
                }
                let visual = if has_table_floats {
                    extent_height(measure)
                } else {
                    table_cell_block_height(block, measure)
                };
                let spacing = match block {
                    LayoutBlock::Paragraph(paragraph) => paragraph
                        .attrs
                        .as_ref()
                        .and_then(|attrs| attrs.spacing.as_ref()),
                    _ => None,
                };
                let before = spacing.and_then(|value| value.before).unwrap_or(0.0);
                let after = spacing.and_then(|value| value.after).unwrap_or(0.0);
                content_height += previous_after.max(before) + visual - before - after;
                previous_after = after;
            }
            if is_rotated(source_cell) {
                content_height = longest_line_width(&measured_cell.blocks);
                previous_after = 0.0;
            }
            let padding_height = source_cell
                .padding
                .as_ref()
                .map_or(DEFAULT_CELL_PADDING_Y, |padding| padding.top)
                + source_cell
                    .padding
                    .as_ref()
                    .map_or(DEFAULT_CELL_PADDING_Y, |padding| padding.bottom);
            measured_cell.height =
                (content_height + previous_after).max(float_bottom) + padding_height;
            if source_cell.row_span.unwrap_or(1.0) <= 1.0 {
                max_height = max_height.max(measured_cell.height);
                max_padding_height = max_padding_height.max(padding_height);
            }
            max_border_height = max_border_height.max(cell_border_height(source_cell));
        }
        exact[row_index] = source_row.is_exact_height();
        measured_row.height = match (source_row.height, source_row.height_rule.as_deref()) {
            (Some(height), Some("exact")) => height,
            (Some(height), _) => max_height.max(height + max_padding_height) + max_border_height,
            (None, _) => max_height + max_border_height,
        };
    }

    let mut spanning_cells = Vec::new();
    for row_index in 0..rows.len() {
        for cell_index in 0..table.rows[row_index].cells.len() {
            let source_cell = &table.rows[row_index].cells[cell_index];
            let row_span = source_cell.row_span.unwrap_or(1.0).max(1.0) as usize;
            if row_span <= 1 {
                continue;
            }
            let last = (row_index + row_span - 1).min(rows.len() - 1);
            let needed = rows[row_index].cells[cell_index].height + cell_border_height(source_cell);
            spanning_cells.push((row_index, last, needed));
        }
    }
    spanning_cells.sort_unstable_by_key(|&(row_index, last, _)| (last, row_index));
    for (row_index, last, needed) in spanning_cells {
        let spanned = rows[row_index..=last]
            .iter()
            .map(|row| row.height)
            .sum::<f64>();
        let deficit = needed - spanned;
        if deficit <= 0.0 {
            continue;
        }
        let mut target = last;
        while target > row_index && exact[target] {
            target -= 1;
        }
        if !exact[target] {
            rows[target].height += deficit;
        }
    }

    for row_index in 0..rows.len() {
        if table.rows[row_index].height.is_none() {
            continue;
        }
        for cell_index in 0..table.rows[row_index].cells.len() {
            let cell = &mut table.rows[row_index].cells[cell_index];
            let row_span = cell.row_span.unwrap_or(1.0).max(1.0) as usize;
            if !is_rotated(cell) || row_span <= 1 {
                continue;
            }
            let last = (row_index + row_span - 1).min(rows.len() - 1);
            let spanned = rows[row_index..=last]
                .iter()
                .map(|row| row.height)
                .sum::<f64>();
            let padding = cell
                .padding
                .as_ref()
                .map_or(0.0, |padding| padding.top + padding.bottom);
            let length = (spanned - padding).max(1.0);
            let has_table_floats = cell.blocks.iter().any(|block| {
                matches!(block, LayoutBlock::Table(table)
                    if nested_table_float_offset(table.floating.as_ref()).is_some())
            });
            let measured = &mut rows[row_index].cells[cell_index];
            measured.blocks = if has_table_floats {
                measure_cell_blocks_with_table_floats(&mut cell.blocks, length, config)?
            } else {
                measure_blocks(&mut cell.blocks, length, config)?
            };
            measured.height = longest_line_width(&measured.blocks) + padding;
            let deficit = measured.height + cell_border_height(cell) - spanned;
            let mut target = last;
            while target > row_index && exact[target] {
                target -= 1;
            }
            if deficit > 0.0 && !exact[target] {
                rows[target].height += deficit;
            }
        }
    }

    let outer_width = |border: &crate::types::CellBorderSpec| {
        if matches!(border.style.as_deref(), Some("none" | "nil")) {
            0.0
        } else {
            border.width.unwrap_or(0.0)
        }
    };
    let mut top_border = 0.0_f64;
    let mut bottom_border = 0.0_f64;
    for cell in &grid {
        let Some(borders) = &table.rows[cell.row_index].cells[cell.cell_index].borders else {
            continue;
        };
        if cell.row_index == 0 {
            top_border = top_border.max(borders.top.as_ref().map_or(0.0, outer_width));
        }
        if cell.row_index + cell.row_span >= rows.len() {
            bottom_border = bottom_border.max(borders.bottom.as_ref().map_or(0.0, outer_width));
        }
    }
    if let Some(row) = rows.first_mut().filter(|_| !exact[0]) {
        row.height += top_border / 2.0;
    }
    if let Some(row) = rows.last_mut() {
        row.height += if exact[exact.len() - 1] {
            bottom_border
        } else {
            bottom_border / 2.0
        };
    }

    let total_height = rows.iter().map(|row| row.height).sum();
    let resolved_total = column_widths.iter().sum::<f64>();
    Ok(TableExtent {
        rows,
        column_widths,
        total_width: if resolved_total != 0.0 {
            resolved_total
        } else {
            explicit_width.unwrap_or(content_width)
        },
        total_height,
    })
}

fn longest_line_width(blocks: &[BlockExtent]) -> f64 {
    blocks
        .iter()
        .filter_map(|measure| match measure {
            BlockExtent::Paragraph(paragraph) => Some(
                paragraph
                    .lines
                    .iter()
                    .map(|line| line.width)
                    .fold(0.0, f64::max),
            ),
            _ => None,
        })
        .fold(0.0, f64::max)
}

fn is_rotated(cell: &crate::types::TableCell) -> bool {
    matches!(cell.text_direction.as_deref(), Some("btLr" | "tbRl"))
}

fn table_cell_block_height(block: &LayoutBlock, measure: &BlockExtent) -> f64 {
    let (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent)) = (block, measure)
    else {
        return extent_height(measure);
    };
    let non_empty: Vec<_> = paragraph
        .runs
        .iter()
        .filter(|run| !matches!(run, Run::Text(text) if text.text.is_empty()))
        .collect();
    let image_only = extent.lines.len() == 1
        && !non_empty.is_empty()
        && non_empty.iter().all(|run| matches!(run, Run::Image(_)));
    if !image_only {
        return extent.total_height;
    }
    let image_height = non_empty
        .iter()
        .filter_map(|run| match run {
            Run::Image(image) => {
                Some(rotation_bound(&image.rotation_bounds, "height").unwrap_or(image.height))
            }
            _ => None,
        })
        .fold(0.0_f64, f64::max);
    let spacing = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.spacing.as_ref());
    // A multiple rule adds its room below an image alone on its line.
    let added = match synthetic_line_rule(spacing) {
        Some(LineSpacingRule::Auto { line_240ths }) if line_240ths > 240 => {
            extent.lines.first().map_or(0.0, |line| {
                (line.line_height - line.ascent - line.descent).max(0.0)
            })
        }
        _ => 0.0,
    };
    spacing.and_then(|value| value.before).unwrap_or(0.0)
        + image_height
        + added
        + spacing.and_then(|value| value.after).unwrap_or(0.0)
}

pub fn extent_height(measure: &BlockExtent) -> f64 {
    match measure {
        BlockExtent::Paragraph(value) => value.total_height,
        BlockExtent::Table(value) => value.total_height,
        BlockExtent::Image(value) => value.height,
        BlockExtent::Shape(value) => value.height,
        BlockExtent::Chart(value) => value.height,
        BlockExtent::TextBox(value) => value.height,
        _ => 0.0,
    }
}

fn cell_border_height(cell: &crate::types::TableCell) -> f64 {
    cell.borders.as_ref().map_or(0.0, |borders| {
        borders
            .top
            .as_ref()
            .and_then(|border| border.width)
            .unwrap_or(0.0)
            + borders
                .bottom
                .as_ref()
                .and_then(|border| border.width)
                .unwrap_or(0.0)
    }) / 2.0
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn cache_measurement_config() -> MeasurementConfig {
        let font = crate::register_measure_font_bytes(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily": "Liberation Sans", "fontSize": 12}),
            ..Default::default()
        }
    }

    fn cache_paragraph(fonts: &crate::MeasureFonts) {
        let _fonts = fonts.enter();
        let config = cache_measurement_config();
        let paragraph = serde_json::from_value(json!({
            "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}]
        }))
        .unwrap();
        let extent = measure_paragraph(&paragraph, 300.0, &config).unwrap();
        assert!(!extent.lines.is_empty());
        assert_ne!(extent.lines[0].synthetic_fallback, Some(true));
    }

    #[test]
    fn distinct_paragraphs_reuse_cached_extents() {
        clear_extent_cache();
        let fonts = crate::MeasureFonts::default();
        let _fonts = fonts.enter();
        let config = cache_measurement_config();
        let paragraphs: Vec<ParagraphBlock> = serde_json::from_value(json!([
            {"id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}]},
            {"id": "cached", "runs": [{"kind": "text", "text": "Changed paragraph"}]},
            {
                "id": "cached",
                "runs": [{"kind": "text", "text": "Cached paragraph", "fontSize": 18}]
            },
            {
                "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}],
                "attrs": {"pPrIns": {"n": 1, "items": ["x", null]}}
            },
            {
                "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}],
                "attrs": {"pPrIns": {"n": 2, "items": ["x", null]}}
            },
            {
                "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}],
                "pmStart": 0.0
            },
            {
                "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}],
                "pmStart": -0.0
            }
        ]))
        .unwrap();
        let first: Vec<_> = paragraphs
            .iter()
            .map(|paragraph| measure_paragraph(paragraph, 300.0, &config).unwrap())
            .collect();
        assert_eq!(extent_cache_stats().0, paragraphs.len());

        for (paragraph, expected) in paragraphs.iter().zip(&first) {
            assert!(!expected.lines.is_empty());
            assert_ne!(expected.lines[0].synthetic_fallback, Some(true));
            let ExtentLookup::Hit(cached) =
                extent_cache_lookup(paragraph, 300.0, &config, None, 0.0)
            else {
                panic!("expected a cached paragraph extent");
            };
            assert_eq!(&cached, expected);
            assert_eq!(
                &measure_paragraph(paragraph, 300.0, &config).unwrap(),
                expected
            );
        }
        assert_eq!(extent_cache_stats().0, paragraphs.len());
    }

    #[test]
    fn appended_fonts_only_invalidate_dependent_extents() {
        let regular = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        let appended = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
        for (label, run, initial, final_chain, expected_calls) in [
            (
                "unused",
                json!({"kind": "text", "text": "Latin", "fontFamily": "Requested"}),
                json!({"requested|0|0": [0]}),
                json!({"unrelated|0|0": [1]}),
                0,
            ),
            (
                "completed chain",
                json!({"kind": "text", "text": "Latin", "fontFamily": "Requested"}),
                json!({"requested|0|0": [0]}),
                json!({"requested|0|0": [0, 1]}),
                1,
            ),
            (
                "changed chain",
                json!({"kind": "text", "text": "Latin", "fontFamily": "Requested"}),
                json!({"requested|0|0": [0]}),
                json!({"requested|0|0": [1]}),
                1,
            ),
            (
                "script slot",
                json!({"kind": "text", "text": "العربية", "fontFamily": "Requested",
                "fontSlots": {"hAnsi": "Requested", "cs": "Script"}}),
                json!({"requested|0|0": [0]}),
                json!({"script|0|0": [1]}),
                1,
            ),
            (
                "missing alternative",
                json!({"kind": "text", "text": "العربية", "fontFamily": "Requested",
                "boldCs": true, "fontSlots": {"hAnsi": "Requested", "cs": "Script"}}),
                json!({"requested|0|0": [0]}),
                json!({"requested|1|0": [1]}),
                1,
            ),
            (
                "missing primary",
                json!({"kind": "text", "text": "Latin", "fontFamily": "Requested"}),
                json!({}),
                json!({"requested|0|0": [1]}),
                1,
            ),
        ] {
            let fonts = crate::MeasureFonts::default();
            let _scope = fonts.enter();
            assert_eq!(crate::register_measure_font_bytes(regular).unwrap(), 0);
            let mut config = MeasurementConfig {
                font_chains: serde_json::from_value(initial).unwrap(),
                defaults: json!({"fontFamily": "Stable", "fontSize": 12}),
                ..Default::default()
            };
            config.font_chains.insert("stable|0|0".to_owned(), vec![0]);
            let paragraphs: Vec<ParagraphBlock> = serde_json::from_value(json!([
                {"id": "dependent", "runs": [run]},
                {"id": "stable", "runs": [{"kind": "text", "text": "Unchanged", "fontFamily": "Stable"}]}
            ])).unwrap();
            for paragraph in &paragraphs {
                measure_paragraph(paragraph, 300.0, &config).unwrap();
            }
            assert_eq!(crate::register_measure_font_bytes(appended).unwrap(), 1);
            config
                .font_chains
                .extend(serde_json::from_value::<BTreeMap<String, Vec<u32>>>(final_chain).unwrap());
            let before = EXTENT_MEASURE_CALLS.with(|calls| calls.get());
            let warm: Vec<_> = paragraphs
                .iter()
                .map(|paragraph| measure_paragraph(paragraph, 300.0, &config).unwrap())
                .collect();
            assert_eq!(
                EXTENT_MEASURE_CALLS.with(|calls| calls.get()) - before,
                expected_calls,
                "{label}"
            );
            let cold = crate::with_private_measure_fonts(|| {
                assert_eq!(crate::register_measure_font_bytes(regular).unwrap(), 0);
                assert_eq!(crate::register_measure_font_bytes(appended).unwrap(), 1);
                paragraphs
                    .iter()
                    .map(|paragraph| measure_paragraph(paragraph, 300.0, &config).unwrap())
                    .collect::<Vec<_>>()
            });
            assert_eq!(
                serde_json::to_vec(&warm).unwrap(),
                serde_json::to_vec(&cold).unwrap(),
                "{label}"
            );
        }
    }

    #[test]
    fn first_font_invalidates_extents_without_a_matching_chain() {
        let fonts = crate::MeasureFonts::default();
        let _scope = fonts.enter();
        let mut config = MeasurementConfig {
            defaults: json!({"fontFamily": "Requested", "fontSize": 12}),
            ..Default::default()
        };
        let paragraph: ParagraphBlock = serde_json::from_value(json!({
            "id": "empty-run", "runs": [
                {"kind": "text", "text": "", "fontFamily": "Requested"},
                {"kind": "lineBreak"}
            ]
        }))
        .unwrap();
        let (empty, dependencies) = FontChainDependencies::capture(|| {
            measure_paragraph(&paragraph, 300.0, &config).unwrap()
        });
        assert_eq!(empty.lines.len(), 2);
        assert_ne!(empty.lines[0].synthetic_fallback, Some(true));
        assert!(matches!(
            extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
            ExtentLookup::Hit(_)
        ));
        let font = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        assert_eq!(crate::register_measure_font_bytes(font).unwrap(), 0);
        config
            .font_chains
            .insert("unrelated|0|0".to_owned(), vec![0]);
        assert!(dependencies.matches(FontChains::BTree(&config.font_chains)));
        assert!(matches!(
            extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
            ExtentLookup::Miss(_)
        ));
        let before = EXTENT_MEASURE_CALLS.with(|calls| calls.get());
        let warm = measure_paragraph(&paragraph, 300.0, &config).unwrap();
        assert_eq!(EXTENT_MEASURE_CALLS.with(|calls| calls.get()) - before, 1);
        assert_ne!(
            serde_json::to_vec(&empty).unwrap(),
            serde_json::to_vec(&warm).unwrap()
        );
        let cold = crate::with_private_measure_fonts(|| {
            assert_eq!(crate::register_measure_font_bytes(font).unwrap(), 0);
            measure_paragraph(&paragraph, 300.0, &config).unwrap()
        });
        assert_eq!(
            serde_json::to_vec(&warm).unwrap(),
            serde_json::to_vec(&cold).unwrap()
        );
    }

    #[test]
    fn registered_chain_font_invalidates_synthetic_extents() {
        let fonts = crate::MeasureFonts::default();
        let _scope = fonts.enter();
        let regular = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        let appended = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
        assert_eq!(crate::register_measure_font_bytes(regular).unwrap(), 0);
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("requested|0|0".to_owned(), vec![1])]),
            defaults: json!({"fontFamily": "Requested", "fontSize": 12}),
            ..Default::default()
        };
        let paragraph: ParagraphBlock = serde_json::from_value(json!({
            "id": "pending-font", "runs": [
                {"kind": "text", "text": "Latin", "fontFamily": "Requested"}
            ]
        }))
        .unwrap();
        let (synthetic, dependencies) = FontChainDependencies::capture(|| {
            measure_paragraph(&paragraph, 300.0, &config).unwrap()
        });
        assert_eq!(synthetic.lines[0].synthetic_fallback, Some(true));
        assert!(matches!(
            extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
            ExtentLookup::Hit(_)
        ));
        assert_eq!(crate::register_measure_font_bytes(appended).unwrap(), 1);
        assert!(dependencies.matches(FontChains::BTree(&config.font_chains)));
        assert!(matches!(
            extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
            ExtentLookup::Miss(_)
        ));
        let before = EXTENT_MEASURE_CALLS.with(|calls| calls.get());
        let warm = measure_paragraph(&paragraph, 300.0, &config).unwrap();
        assert_eq!(EXTENT_MEASURE_CALLS.with(|calls| calls.get()) - before, 1);
        assert_ne!(warm.lines[0].synthetic_fallback, Some(true));
        assert_ne!(
            serde_json::to_vec(&synthetic).unwrap(),
            serde_json::to_vec(&warm).unwrap()
        );
        let cold = crate::with_private_measure_fonts(|| {
            assert_eq!(crate::register_measure_font_bytes(regular).unwrap(), 0);
            assert_eq!(crate::register_measure_font_bytes(appended).unwrap(), 1);
            measure_paragraph(&paragraph, 300.0, &config).unwrap()
        });
        assert_eq!(
            serde_json::to_vec(&warm).unwrap(),
            serde_json::to_vec(&cold).unwrap()
        );
        assert_eq!(crate::register_measure_font_bytes(regular).unwrap(), 2);
        let before = EXTENT_MEASURE_CALLS.with(|calls| calls.get());
        assert_eq!(measure_paragraph(&paragraph, 300.0, &config).unwrap(), warm);
        assert_eq!(EXTENT_MEASURE_CALLS.with(|calls| calls.get()), before);
    }

    #[test]
    fn extent_dependencies_cover_marks_markers_tabs_fields_and_caps() {
        let fonts = crate::MeasureFonts::default();
        let _scope = fonts.enter();
        let mut config = cache_measurement_config();
        config
            .font_chains
            .insert("requested|0|0".to_owned(), vec![0]);
        let appended = crate::register_measure_font_bytes(include_bytes!(
            "../../docx-raster/tests/assets/Carlito-Regular.ttf"
        ))
        .unwrap();
        for (label, runs, attrs) in [
            (
                "paragraph mark",
                json!([]),
                json!({"defaultFontFamily": "Requested"}),
            ),
            (
                "marker",
                json!([{"kind": "text", "text": "Body"}]),
                json!({"listMarker": "1.", "listMarkerFontFamily": "Requested"}),
            ),
            (
                "tab",
                json!([{"kind": "tab", "fontFamily": "Requested"}]),
                json!(null),
            ),
            (
                "field",
                json!([{"kind": "field", "fieldType": "PAGE", "fallback": "1", "fontFamily": "Requested"}]),
                json!(null),
            ),
            (
                "caps",
                json!([{"kind": "text", "text": "ßabc", "allCaps": true, "fontFamily": "Requested"}]),
                json!(null),
            ),
            (
                "small caps",
                json!([{"kind": "text", "text": "abc", "smallCaps": true, "fontFamily": "Requested"}]),
                json!(null),
            ),
        ] {
            config
                .font_chains
                .insert("requested|0|0".to_owned(), vec![0]);
            let paragraph: ParagraphBlock = serde_json::from_value(json!({
                "id": label, "runs": runs, "attrs": attrs
            }))
            .unwrap();
            let (extent, dependencies) = FontChainDependencies::capture(|| {
                measure_paragraph(&paragraph, 300.0, &config).unwrap()
            });
            assert_ne!(extent.lines[0].synthetic_fallback, Some(true), "{label}");
            config
                .font_chains
                .insert("requested|0|0".to_owned(), vec![appended]);
            assert!(
                !dependencies.matches(FontChains::BTree(&config.font_chains)),
                "{label}"
            );
            assert!(
                matches!(
                    extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
                    ExtentLookup::Miss(_)
                ),
                "{label}"
            );
        }
        let paragraph: ParagraphBlock = serde_json::from_value(json!({
            "id": "minimum", "runs": [{"kind": "text", "text": "unbreakableword", "fontFamily": "Requested"}]
        })).unwrap();
        let (width, dependencies) = FontChainDependencies::capture(|| {
            crate::typed_measure::min_content_width(&paragraph, 300.0, &config)
        });
        assert!(width.is_some());
        config.font_chains.remove("requested|0|0");
        assert!(!dependencies.matches(FontChains::BTree(&config.font_chains)));
    }

    #[test]
    fn extent_dependencies_survive_cache_hits_and_store_changes() {
        let fonts = crate::MeasureFonts::default();
        let _scope = fonts.enter();
        let mut config = cache_measurement_config();
        let paragraph: ParagraphBlock = serde_json::from_value(json!({
            "id": "cached", "runs": [{"kind": "text", "text": "Cached paragraph"}]
        }))
        .unwrap();
        measure_paragraph(&paragraph, 300.0, &config).unwrap();
        let (_, dependencies) = FontChainDependencies::capture(|| {
            measure_paragraph(&paragraph, 300.0, &config).unwrap()
        });
        config.font_chains.remove("liberation sans|0|0");
        assert!(!dependencies.matches(FontChains::BTree(&config.font_chains)));
        config = cache_measurement_config();
        assert!(matches!(
            extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
            ExtentLookup::Miss(_)
        ));
        measure_paragraph(&paragraph, 300.0, &config).unwrap();
        crate::with_private_measure_fonts(|| {
            crate::register_measure_font_bytes(include_bytes!(
                "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
            ))
            .unwrap();
            crate::register_measure_font_bytes(include_bytes!(
                "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
            ))
            .unwrap();
            assert!(matches!(
                extent_cache_lookup(&paragraph, 300.0, &config, None, 0.0),
                ExtentLookup::Miss(_)
            ));
        });
    }

    #[test]
    fn dropping_last_measure_fonts_releases_extent_cache_and_scratch() {
        clear_extent_cache();
        let fonts = crate::MeasureFonts::default();
        cache_paragraph(&fonts);
        let (entries, scratch_capacity) = extent_cache_stats();
        assert_eq!(entries, 1);
        assert!(scratch_capacity > 0);

        drop(fonts);
        assert_eq!(extent_cache_stats(), (0, 0));
    }

    #[test]
    fn dropping_measure_fonts_keeps_extent_cache_while_another_is_alive() {
        clear_extent_cache();
        let (a, b) = (
            crate::MeasureFonts::default(),
            crate::MeasureFonts::default(),
        );
        cache_paragraph(&a);
        cache_paragraph(&b);
        let populated = extent_cache_stats();
        assert_eq!(populated.0, 2);
        assert!(populated.1 > 0);

        drop(a);
        assert_eq!(extent_cache_stats(), populated);

        drop(b);
        assert_eq!(extent_cache_stats(), (0, 0));
    }

    #[test]
    fn wide_floating_tables_with_no_room_on_their_side_wrap_in_the_larger_gap() {
        let measure = TableExtent {
            rows: Vec::new(),
            column_widths: vec![360.0],
            total_width: 360.0,
            total_height: 160.0,
        };
        for (x, left_distance, right_distance, expected) in [
            (0.0, 9.0, 13.0, (373.0, 0.0)),
            (240.0, 9.0, 13.0, (0.0, 369.0)),
            (100.0, 9.0, 13.0, (473.0, 0.0)),
            (140.0, 9.0, 13.0, (513.0, 0.0)),
            (120.0, 9.0, 13.0, (493.0, 0.0)),
            (120.0, 9.0, 9.0, (489.0, 0.0)),
            (230.0, 9.0, 13.0, (0.0, 379.0)),
        ] {
            let table = serde_json::from_value(json!({
                "id": "float", "rows": [],
                "floating": {
                    "horzAnchor": "text", "tblpX": x,
                    "leftFromText": left_distance, "rightFromText": right_distance
                }
            }))
            .unwrap();
            let zone = table_floating_zone(&table, &measure, 600.0, Some(600.0)).unwrap();
            assert_eq!((zone.left_margin, zone.right_margin), expected, "x={x}");
        }
        for (spec, expected) in [
            ("left", (373.0, 0.0)),
            ("right", (0.0, 369.0)),
            ("center", (493.0, 0.0)),
        ] {
            let table = serde_json::from_value(json!({
                "id": "float", "rows": [],
                "floating": {
                    "horzAnchor": "text", "tblpXSpec": spec,
                    "leftFromText": 9, "rightFromText": 13
                }
            }))
            .unwrap();
            let zone = table_floating_zone(&table, &measure, 600.0, Some(600.0)).unwrap();
            assert_eq!((zone.left_margin, zone.right_margin), expected, "{spec}");
        }
    }

    #[test]
    fn narrow_floating_tables_keep_their_existing_margins() {
        for width in [80.0, 299.0, 300.0] {
            let measure = TableExtent {
                rows: Vec::new(),
                column_widths: vec![width],
                total_width: width,
                total_height: 160.0,
            };
            for x in [0.0, 120.0, 299.0, 300.0, 600.0 - width] {
                let table = serde_json::from_value(json!({
                    "id": "float", "rows": [],
                    "floating": {
                        "horzAnchor": "text", "tblpX": x,
                        "leftFromText": 9.4, "rightFromText": 13.2
                    }
                }))
                .unwrap();
                let zone = table_floating_zone(&table, &measure, 600.0, Some(600.0)).unwrap();
                let expected: (f64, f64) = if x < 300.0 {
                    (x + width + 13.2, 0.0)
                } else {
                    (0.0, 600.0 - x + 9.4)
                };
                assert_eq!(zone.left_margin.to_bits(), expected.0.to_bits());
                assert_eq!(zone.right_margin.to_bits(), expected.1.to_bits());
            }
        }
    }

    #[test]
    fn floating_tables_with_two_small_gaps_keep_their_margins() {
        let measure = TableExtent {
            rows: Vec::new(),
            column_widths: vec![540.0],
            total_width: 540.0,
            total_height: 160.0,
        };
        for (floating, expected) in [
            (
                json!({"horzAnchor": "text", "tblpXSpec": "center"}),
                (582.0, 0.0),
            ),
            (
                json!({"horzAnchor": "text", "tblpXSpec": "center", "leftFromText": 7, "rightFromText": 7}),
                (577.0, 0.0),
            ),
        ] {
            let table = serde_json::from_value(json!({
                "id": "float", "rows": [], "floating": floating
            }))
            .unwrap();
            let zone = table_floating_zone(&table, &measure, 600.0, Some(600.0)).unwrap();
            assert_eq!((zone.left_margin, zone.right_margin), expected);
        }
    }

    #[test]
    fn wide_floating_tables_outside_the_column_frame_keep_their_margins() {
        let measure = TableExtent {
            rows: Vec::new(),
            column_widths: vec![360.0],
            total_width: 360.0,
            total_height: 160.0,
        };
        for (anchor, column_width, expected) in [
            (Some("page"), Some(600.0), (522.0, 0.0)),
            (None, Some(600.0), (522.0, 0.0)),
            (Some("margin"), Some(600.0), (522.0, 0.0)),
            (Some("margin"), None, (522.0, 0.0)),
            (Some("text"), Some(600.0), (522.0, 0.0)),
            (Some("column"), Some(600.0), (522.0, 0.0)),
            (Some("text"), Some(800.0), (522.0, 0.0)),
            (Some("text"), None, (522.0, 0.0)),
        ] {
            let table = serde_json::from_value(json!({
                "id": "float", "rows": [],
                "floating": {
                    "horzAnchor": anchor, "tblpX": 150, "leftFromText": 12, "rightFromText": 12
                }
            }))
            .unwrap();
            let zone = table_floating_zone(&table, &measure, 600.0, column_width).unwrap();
            assert_eq!(
                (zone.left_margin, zone.right_margin),
                expected,
                "{anchor:?} {column_width:?}"
            );
        }
    }

    #[test]
    fn wide_floating_tables_with_unknown_section_frames_keep_main_margins() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
            "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
            "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
            "floating": {
                "horzAnchor": "text", "tblpXSpec": "center", "leftFromText": 9, "rightFromText": 13
            }
        }]))
        .unwrap();
        let flow = FloatFlow::new(&blocks, &[600.0], &MeasurementConfig::default(), None).unwrap();
        let zone = &flow.paragraph_zones[&0][0];
        assert_eq!((zone.left_margin, zone.right_margin), (493.0, 0.0));
    }

    #[test]
    fn wide_floating_tables_beside_negative_indents_keep_main_margins() {
        for (indent, expected) in [
            (json!({}), (0.0, 369.0)),
            (json!({"right": -100}), (0.0, 0.0)),
            (json!({"left": 20, "hanging": 40}), (0.0, 0.0)),
        ] {
            let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {
                    "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
                    "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                    "floating": {
                        "horzAnchor": "text", "tblpX": 240, "leftFromText": 9, "rightFromText": 13
                    }
                },
                {"kind": "paragraph", "id": "text", "attrs": {"indent": indent}, "runs": []},
                {"kind": "pageBreak", "id": "break"},
                {"kind": "paragraph", "id": "later", "attrs": {"indent": {"right": -100}}, "runs": []}
            ]))
            .unwrap();
            let frames: Vec<bool> = negative_indent_float_flows(&blocks.iter().collect::<Vec<_>>())
                .into_iter()
                .map(|negative| !negative)
                .collect();
            assert_eq!(frames[2..], [false, false]);
            let flow = FloatFlow::with_table_wrap_frames(
                &blocks,
                &[600.0; 4],
                &frames,
                &MeasurementConfig::default(),
                None,
            )
            .unwrap();
            let zone = &flow.paragraph_zones[&0][0];
            assert_eq!((zone.left_margin, zone.right_margin), expected);
        }
    }

    #[test]
    fn cell_table_floats_keep_main_wrap_sides() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily": "Liberation Sans", "fontSize": 12}),
            ..Default::default()
        };
        for (anchor, expected) in [("text", (493.0, 0.0)), ("margin", (493.0, 0.0))] {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {
                    "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
                    "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                    "floating": {
                        "horzAnchor": anchor, "tblpXSpec": "center", "vertAnchor": "text", "tblpY": 1,
                        "leftFromText": 9, "rightFromText": 13
                    }
                },
                {"kind": "paragraph", "id": "text", "runs": [{"kind": "text", "text": "short line"}]}
            ]))
            .unwrap();
            let measures =
                measure_cell_blocks_with_table_floats(&mut blocks, 600.0, &config).unwrap();
            let BlockExtent::Paragraph(paragraph) = &measures[1] else {
                panic!()
            };
            let line = &paragraph.lines[0];
            assert_ne!(line.synthetic_fallback, Some(true));
            assert_eq!(
                (
                    line.left_offset.unwrap_or(0.0),
                    line.right_offset.unwrap_or(0.0)
                ),
                expected,
                "{anchor}"
            );
        }
    }

    #[test]
    fn wide_floating_tables_with_unstable_positions_keep_their_margins() {
        let measure = TableExtent {
            rows: Vec::new(),
            column_widths: vec![360.0],
            total_width: 360.0,
            total_height: 100.0,
        };
        for (position, expected) in [
            (json!({"tblpXSpec": "outside"}), (613.0_f64, 0.0_f64)),
            (json!({"tblpXSpec": "inside"}), (373.0, 0.0)),
            (json!({"tblpXSpec": "outside", "tblpX": 150}), (523.0, 0.0)),
            (json!({"tblpXSpec": "inside", "tblpX": 150}), (523.0, 0.0)),
            (json!({}), (493.0, 0.0)),
            (json!({"tblpXSpec": "unknown"}), (493.0, 0.0)),
        ] {
            let mut floating = json!({
                "horzAnchor": "text", "leftFromText": 9, "rightFromText": 13
            });
            floating
                .as_object_mut()
                .unwrap()
                .extend(position.as_object().unwrap().clone());
            let table = serde_json::from_value(json!({
                "id": "float", "rows": [], "justification": "center", "floating": floating
            }))
            .unwrap();
            let zone = table_floating_zone(&table, &measure, 600.0, Some(600.0)).unwrap();
            assert_eq!(
                zone.left_margin.to_bits(),
                expected.0.to_bits(),
                "{position}"
            );
            assert_eq!(
                zone.right_margin.to_bits(),
                expected.1.to_bits(),
                "{position}"
            );
        }
    }

    #[test]
    fn margin_anchored_tables_in_unequal_columns_keep_their_margins_without_geometry_too() {
        let geometry = FloatPageGeometry {
            page_width: 800.0,
            margin_left: 100.0,
            page_height: 800.0,
            margin_top: 100.0,
            content_height: 600.0,
        };
        for (column_width, table_width, expected) in [
            (360.0, 200.0, (293.0_f64, 0.0_f64)),
            (280.0, 180.0, (243.0, 0.0)),
        ] {
            let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind": "columnBreak", "id": "break"},
                {
                    "kind": "table", "id": "float", "columnWidths": [table_width], "layoutMode": "fixed",
                    "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                    "floating": {
                        "horzAnchor": "margin", "tblpXSpec": "center",
                        "vertAnchor": "text", "tblpY": 1,
                        "leftFromText": 9, "rightFromText": 13
                    }
                },
                {"kind": "paragraph", "id": "text", "runs": [{"kind": "text", "text": "test"}]}
            ]))
            .unwrap();
            for page_geometry in [Some(&geometry), None] {
                let flow = FloatFlow::new(
                    &blocks,
                    &[column_width; 3],
                    &MeasurementConfig::default(),
                    page_geometry,
                )
                .unwrap();
                let zone = &flow.paragraph_zones[&1][0];
                assert_eq!(zone.left_margin.to_bits(), expected.0.to_bits());
                assert_eq!(zone.right_margin.to_bits(), expected.1.to_bits());
            }
        }
    }

    #[test]
    fn outside_anchored_tables_on_even_pages_keep_their_margins() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {"kind": "pageBreak", "id": "page-two"},
            {
                "kind": "table", "id": "float", "columnWidths": [360], "layoutMode": "fixed",
                "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                "floating": {
                    "horzAnchor": "text", "tblpXSpec": "outside",
                    "vertAnchor": "text", "tblpY": 1,
                    "leftFromText": 9, "rightFromText": 13
                }
            },
            {
                "kind": "paragraph", "id": "text", "attrs": {"alignment": "right"},
                "runs": [{"kind": "text", "text": "test"}]
            }
        ]))
        .unwrap();
        let geometry = FloatPageGeometry {
            page_width: 800.0,
            margin_left: 100.0,
            page_height: 800.0,
            margin_top: 100.0,
            content_height: 600.0,
        };
        let flow = FloatFlow::new(
            &blocks,
            &[600.0; 3],
            &MeasurementConfig::default(),
            Some(&geometry),
        )
        .unwrap();
        let zone = &flow.paragraph_zones[&1][0];
        assert_eq!(zone.left_margin.to_bits(), 0.0_f64.to_bits());
        assert_eq!(zone.right_margin.to_bits(), 0.0_f64.to_bits());
    }

    #[test]
    fn later_section_tables_with_a_different_column_width_keep_their_margins() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {"kind": "paragraph", "id": "first-section", "runs": []},
            {
                "kind": "sectionBreak", "id": "second-section", "type": "nextPage",
                "pageSize": {"w": 1000, "h": 800},
                "margins": {"top": 100, "right": 100, "bottom": 100, "left": 100}
            },
            {
                "kind": "table", "id": "float", "columnWidths": [300], "layoutMode": "fixed",
                "rows": [{"id": "row", "height": 100, "heightRule": "exact", "cells": []}],
                "floating": {
                    "horzAnchor": "text", "tblpXSpec": "center",
                    "vertAnchor": "text", "tblpY": 1,
                    "leftFromText": 9, "rightFromText": 13
                }
            },
            {
                "kind": "paragraph", "id": "text", "attrs": {"alignment": "right"},
                "runs": [{"kind": "text", "text": "test"}]
            }
        ]))
        .unwrap();
        let widths = [400.0, 400.0, 800.0, 800.0];
        let geometry = FloatPageGeometry {
            page_width: 600.0,
            margin_left: 100.0,
            page_height: 800.0,
            margin_top: 100.0,
            content_height: 600.0,
        };
        let config = MeasurementConfig::default();
        let flow = FloatFlow::new(&blocks, &widths, &config, Some(&geometry)).unwrap();
        let zone = &flow.paragraph_zones[&2][0];
        assert_eq!(zone.left_margin.to_bits(), 363.0_f64.to_bits());
        assert_eq!(zone.right_margin.to_bits(), 0.0_f64.to_bits());
        let segment = extract_floating_zones(
            &blocks[1..],
            400.0,
            &widths[1..],
            &[],
            &config,
            Some(&geometry),
            &BTreeMap::new(),
        )
        .unwrap();
        assert_eq!(
            segment[0].zone.left_margin.to_bits(),
            zone.left_margin.to_bits()
        );
        assert_eq!(
            segment[0].zone.right_margin.to_bits(),
            zone.right_margin.to_bits()
        );
    }

    #[test]
    fn a_paragraph_anchored_band_hangs_off_its_own_anchor_not_an_earlier_one() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let shape = |id: &str, offset: f64| {
            json!({"kind":"shape","id":id,"shapeType":"rect","geometryPath":[],"children":[],
                "width":40,"height":40,"wrapType":"square",
                "wrapDistances":{"left":0,"right":10,"top":0,"bottom":0},
                "position":{"horizontal":{"relativeTo":"column","posOffset":0},
                    "vertical":{"relativeTo":"paragraph","posOffset":offset}}})
        };
        let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            shape("far", 30.0),
            {"kind":"paragraph","id":"first","runs":[{"kind":"text","text":"words"}]},
            shape("near", 0.0),
            {"kind":"paragraph","id":"second","runs":[{"kind":"text","text":"words"}]}
        ]))
        .unwrap();
        let measures = measure_blocks_with_floats(&mut blocks, &[200.0; 4], &config, None).unwrap();
        let offset = |measure: &BlockExtent| {
            let BlockExtent::Paragraph(paragraph) = measure else {
                panic!()
            };
            paragraph.lines[0].left_offset.unwrap_or(0.0)
        };
        // `near` starts at the second paragraph, below the first one's line.
        assert_eq!(offset(&measures[1]), 0.0);
        assert_eq!(offset(&measures[3]), 50.0);
    }

    /// Measured at 25px on `oxi-en-correspondence-03` at 150dpi: Word gives a
    /// section holding only its break mark one line.
    #[test]
    fn a_section_break_mark_keeps_its_line_when_it_is_all_the_section_holds() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let measure = |blocks: serde_json::Value| {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(blocks).unwrap();
            measure_blocks_with_floats(&mut blocks, &[200.0; 6], &config, None).unwrap()
        };
        let empty = json!({"kind":"paragraph","id":"mark","runs":[]});
        let spaced = json!({"kind":"paragraph","id":"mark","runs":[{"kind":"text","text":" "}]});
        let section = json!({"kind":"sectionBreak","id":"sect:mark"});
        let lead = json!({"kind":"paragraph","id":"lead","runs":[{"kind":"text","text":"words"}]});
        let tail = json!({"kind":"paragraph","id":"tail","runs":[{"kind":"text","text":"words"}]});
        let mark_of = |measures: &[BlockExtent], index: usize| {
            let BlockExtent::Paragraph(mark) = &measures[index] else {
                panic!()
            };
            (mark.lines.len(), mark.total_height)
        };

        let marked = measure(json!([lead, empty, section, tail]));
        assert_eq!(mark_of(&marked, 1), (0, 0.0));

        for (measures, index) in [
            (measure(json!([empty, section, tail])), 0),
            (measure(json!([lead, section, empty, section, tail])), 2),
            (measure(json!([empty, tail])), 0),
            (measure(json!([lead, spaced, section, tail])), 1),
        ] {
            let (lines, height) = mark_of(&measures, index);
            assert_eq!(lines, 1);
            assert!(height > 0.0);
        }
    }

    #[test]
    fn nested_text_anchored_tables_share_measure_paint_and_break_positions() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..Default::default()
        };
        let paragraph = |text: &str, before: f64, after: f64| {
            json!({"kind":"paragraph","id":text,
                "attrs":{"spacing":{"before":before,"after":after,"line":16,"lineRule":"exact"}},
                "runs":[{"kind":"text","text":text}]})
        };
        let padding = json!({"top":0,"bottom":0,"left":0,"right":0});
        let mut cases = Vec::new();
        for (anchor, offset, width, before, table_top, after_top) in [
            ("text", 0.0, 220.0, 6.0, 24.0, 70.0),
            ("text", 20.0, 220.0, 6.0, 44.0, 90.0),
            ("text", 40.0, 220.0, 6.0, 64.0, 30.0),
            ("text", 20.0, 220.0, 40.0, 44.0, 124.0),
            ("text", 20.0, 80.0, 6.0, 44.0, 30.0),
            ("margin", 20.0, 220.0, 6.0, 24.0, 70.0),
        ] {
            cases.push((
                json!({"vertAnchor":anchor,"horzAnchor":"margin","tblpX":0,"tblpY":offset}),
                width,
                before,
                table_top,
                after_top,
                None,
                None,
                0.0,
                None,
            ));
        }
        for (horizontal, indent, justification, table_x, text_x) in [
            (json!({}), None, None, 0.0, 92.0),
            (json!({}), Some(30.0), None, 30.0, 122.0),
            (json!({}), None, Some("center"), 70.0, 162.0),
            (json!({}), None, Some("right"), 140.0, 0.0),
            (json!({"tblpXSpec":"left"}), None, None, 0.0, 92.0),
            (json!({"tblpXSpec":"center"}), None, None, 70.0, 162.0),
            (json!({"tblpXSpec":"right"}), None, None, 140.0, 0.0),
            (
                json!({"tblpX":140,"tblpXSpec":"left"}),
                None,
                None,
                0.0,
                92.0,
            ),
            (
                json!({"tblpX":140,"tblpXSpec":"center"}),
                None,
                None,
                70.0,
                162.0,
            ),
            (
                json!({"tblpX":0,"tblpXSpec":"right"}),
                None,
                None,
                140.0,
                0.0,
            ),
        ] {
            let mut floating = json!({"vertAnchor":"text","horzAnchor":"margin","tblpY":20});
            floating
                .as_object_mut()
                .unwrap()
                .extend(horizontal.as_object().unwrap().clone());
            cases.push((
                floating,
                80.0,
                6.0,
                44.0,
                30.0,
                indent,
                justification,
                table_x,
                Some(text_x),
            ));
        }
        for (
            floating,
            width,
            before,
            table_top,
            after_top,
            indent,
            justification,
            table_x,
            text_x,
        ) in cases
        {
            let mut block: LayoutBlock = serde_json::from_value(json!({
                "kind":"table","id":"outer","columnWidths":[220],
                "rows":[{"id":"outer-row","cells":[{"id":"outer-cell","padding":padding,"blocks":[
                    paragraph("Before",0.0,8.0),
                    {"kind":"table","id":"nested",
                     "columnWidths":[width],"floating":floating,"indent":indent,"justification":justification,
                     "rows":[{"id":"nested-row","height":40,"heightRule":"exact","cells":[{"id":"nested-cell","padding":padding,"blocks":[paragraph("Table",0.0,0.0)]}]}]},
                    paragraph("After",before,0.0),
                    paragraph("Tail",0.0,0.0)
                ]}]}]
            }))
            .unwrap();
            let measure = measure_block(&mut block, 220.0, &config).unwrap();
            let (LayoutBlock::Table(table), BlockExtent::Table(extent)) = (&block, &measure) else {
                panic!()
            };
            let cell = &table.rows[0].cells[0];
            let measured_cell = &extent.rows[0].cells[0];
            let layout = crate::cell_layout::layout_cell_content(
                Some(&cell.blocks),
                Some(&measured_cell.blocks),
                0.0,
            );
            assert_eq!(layout.line_tops[2], vec![after_top], "{floating} {width}");
            assert_eq!(layout.line_tops[3], vec![after_top + 16.0]);
            let expected_height = (after_top + 32.0).max(table_top + 40.0);
            assert_eq!(extent.total_height, expected_height);
            assert_eq!(layout.content_height, expected_height);
            let breaks = crate::table_row_break::build_table_row_break_info(table, extent);
            assert!(
                breaks.break_offsets[0]
                    .iter()
                    .all(|offset| { *offset <= table_top || *offset >= table_top + 40.0 })
            );
            let mut input = crate::types::Input {
                measured: vec![crate::types::MeasuredBlock { block, measure }],
                options: serde_json::from_value(json!({"pageSize":{"w":400,"h":300},
                    "margins":{"top":0,"bottom":0,"left":0,"right":0}}))
                .unwrap(),
            };
            let pages = crate::compute_layout_input(&mut input).unwrap();
            let display = crate::build_display_list(&input, &pages).unwrap();
            let baselines: BTreeMap<String, f64> = display.pages[0]
                .primitives
                .iter()
                .filter_map(|primitive| match primitive {
                    crate::display_list::Primitive::Text(text) => {
                        Some((text.text.clone(), text.baseline_y.as_f64().unwrap()))
                    }
                    crate::display_list::Primitive::GlyphRun(text) => {
                        Some((text.text.clone(), text.glyphs[0].y))
                    }
                    _ => None,
                })
                .collect();
            assert!((baselines["After"] - baselines["Before"] - after_top).abs() < 0.01);
            assert!((baselines["Table"] - baselines["Before"] - table_top).abs() < 0.01);
            let positions: BTreeMap<String, f64> = display.pages[0]
                .primitives
                .iter()
                .filter_map(|primitive| match primitive {
                    crate::display_list::Primitive::Text(text) => {
                        Some((text.text.clone(), text.x.as_f64().unwrap()))
                    }
                    crate::display_list::Primitive::GlyphRun(text) => {
                        Some((text.text.clone(), text.glyphs[0].x))
                    }
                    _ => None,
                })
                .collect();
            assert!(
                (positions["Table"] - positions["Before"] - table_x).abs() < 0.01,
                "{floating}"
            );
            if let Some(text_x) = text_x {
                assert!(
                    (positions["After"] - positions["Before"] - text_x).abs() < 0.01,
                    "{floating}"
                );
            }

            input.options.page_size.as_mut().unwrap().h = 80.0;
            let split_pages = crate::compute_layout_input(&mut input).unwrap();
            assert!(split_pages.pages.len() > 1);
            let split_display = crate::build_display_list(&input, &split_pages).unwrap();
            let mut painted = BTreeMap::<String, usize>::new();
            for (page, display) in split_pages.pages.iter().zip(&split_display.pages) {
                let crate::types::Fragment::Table(fragment) = &page.fragments[0] else {
                    panic!()
                };
                let start = fragment.clip_top.unwrap_or(0.0);
                let end = start + fragment.height;
                assert!(start <= table_top || start >= table_top + 40.0);
                assert!(end <= table_top || end >= table_top + 40.0);
                for primitive in &display.primitives {
                    let (text, baseline) = match primitive {
                        crate::display_list::Primitive::Text(text) => {
                            (&text.text, text.baseline_y.as_f64().unwrap())
                        }
                        crate::display_list::Primitive::GlyphRun(text) => {
                            (&text.text, text.glyphs[0].y)
                        }
                        _ => continue,
                    };
                    assert!((baseline + start - baselines[text]).abs() < 0.01);
                    if baseline >= fragment.y && baseline < fragment.y + fragment.height {
                        *painted.entry(text.clone()).or_default() += 1;
                    }
                }
            }
            assert_eq!(painted.len(), 4);
            assert!(
                painted.values().all(|count| *count == 1),
                "{floating} {width}: {painted:?}"
            );
        }
    }

    #[test]
    fn consecutive_nested_floats_keep_anchor_spacing_and_atomic_row_splits() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..Default::default()
        };
        let paragraph = |text: &str, before: f64, after: f64| {
            json!({"kind":"paragraph","id":text,
                "attrs":{"spacing":{"before":before,"after":after,"line":16,"lineRule":"exact"}},
                "runs":[{"kind":"text","text":text}]})
        };
        let padding = json!({"top":0,"bottom":0,"left":0,"right":0});
        let nested = |offset: f64| {
            json!({"kind":"table","id":offset,"columnWidths":[220],
                "floating":{"vertAnchor":"text","horzAnchor":"margin","tblpX":0,"tblpY":offset},
                "rows":[{"id":offset,"height":40,"heightRule":"exact","cells":[
                    {"id":offset,"padding":padding,"blocks":[paragraph("Table",0.0,0.0)]}]}]})
        };
        for offsets in [[20.0, 80.0], [80.0, 20.0]] {
            let mut block: LayoutBlock = serde_json::from_value(json!({
                "kind":"table","id":"outer","columnWidths":[220],
                "rows":[{"id":"row","cells":[{"id":"cell","padding":padding,"blocks":[
                    paragraph("Before",0.0,8.0), nested(offsets[0]), nested(offsets[1]),
                    paragraph("After",6.0,0.0), paragraph("Tail",0.0,0.0)
                ]}]}]
            }))
            .unwrap();
            let measure = measure_block(&mut block, 220.0, &config).unwrap();
            let (LayoutBlock::Table(table), BlockExtent::Table(extent)) = (&block, &measure) else {
                panic!()
            };
            let layout = crate::cell_layout::layout_cell_content(
                Some(&table.rows[0].cells[0].blocks),
                Some(&extent.rows[0].cells[0].blocks),
                0.0,
            );
            assert_eq!(layout.line_tops[3], vec![150.0]);
            assert_eq!(layout.line_tops[4], vec![166.0]);
            assert_eq!(layout.content_height, 182.0);
            assert_eq!(extent.total_height, 182.0);
            let info = crate::table_row_break::build_table_row_break_info(table, extent);
            assert_eq!(info.break_offsets[0], vec![16.0, 84.0, 144.0, 166.0, 182.0]);
            assert_eq!(
                crate::table_row_break::snap_row_break(&info, 0, 0.0, 80.0),
                16.0
            );
            assert_eq!(
                crate::table_row_break::snap_row_break(&info, 0, 16.0, 80.0),
                68.0
            );
            assert_eq!(
                crate::table_row_break::snap_row_break(&info, 0, 84.0, 80.0),
                60.0
            );
        }

        let mut image_block: LayoutBlock = serde_json::from_value(json!({
            "kind":"table","id":"outer","columnWidths":[220],
            "rows":[{"id":"row","cells":[{"id":"cell","padding":padding,"blocks":[
                paragraph("Before",0.0,8.0), nested(20.0),
                {"kind":"paragraph","id":"image","attrs":{"spacing":{"before":6}},
                 "runs":[{"kind":"image","src":"logo","width":20,"height":20}]},
                paragraph("Tail",0.0,0.0)
            ]}]}]
        }))
        .unwrap();
        let image_measure = measure_block(&mut image_block, 220.0, &config).unwrap();
        let (LayoutBlock::Table(table), BlockExtent::Table(extent)) =
            (&image_block, &image_measure)
        else {
            panic!()
        };
        let layout = crate::cell_layout::layout_cell_content(
            Some(&table.rows[0].cells[0].blocks),
            Some(&extent.rows[0].cells[0].blocks),
            0.0,
        );
        assert_eq!(layout.line_tops[2], vec![90.0]);
        assert!((layout.content_height - extent.total_height).abs() < 0.01);
        assert!((layout.line_tops[3][0] + 16.0 - extent.total_height).abs() < 0.01);
    }

    #[test]
    fn anchored_shapes_exclude_body_text_without_advancing_the_cursor() {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        for (wrap, expected_offset, expected_skip) in [
            ("square", 50.0, 0.0),
            ("inFront", 0.0, 0.0),
            ("topAndBottom", 0.0, 40.0),
        ] {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind":"shape","id":"shape","shapeType":"rect","geometryPath":[],"children":[],
                 "width":40,"height":40,"wrapType":wrap,"wrapDistances":{"left":0,"right":10,"top":0,"bottom":0},
                 "position":{"horizontal":{"relativeTo":"column","posOffset":0},"vertical":{"relativeTo":"paragraph","posOffset":0}}},
                {"kind":"paragraph","id":"body","runs":[{"kind":"text","text":"words words words words words words words words"}]}
            ])).unwrap();
            let measures =
                measure_blocks_with_floats(&mut blocks, &[200.0, 200.0], &config, None).unwrap();
            let BlockExtent::Paragraph(paragraph) = &measures[1] else {
                panic!()
            };
            assert_eq!(
                paragraph.lines[0].left_offset.unwrap_or(0.0),
                expected_offset
            );
            assert_eq!(
                paragraph.lines[0].float_skip_before.unwrap_or(0.0),
                expected_skip
            );

            blocks.insert(
                1,
                serde_json::from_value(json!({"kind":"pageBreak","id":"break"})).unwrap(),
            );
            let measures =
                measure_blocks_with_floats(&mut blocks, &[200.0; 3], &config, None).unwrap();
            let BlockExtent::Paragraph(paragraph) = &measures[2] else {
                panic!()
            };
            assert_eq!(paragraph.lines[0].left_offset.unwrap_or(0.0), 0.0);
            assert_eq!(paragraph.lines[0].float_skip_before.unwrap_or(0.0), 0.0);
        }
    }

    /// Word 16.113, probed at twip resolution: the box a paragraph's first
    /// line is tested against runs from the paragraph top through the line's
    /// bottom, so it covers space-before. A `topAndBottom` band reaching into
    /// that box — including the band the paragraph anchors itself — moves the
    /// line below the band and spends space-before again underneath it, at
    /// every band height probed from 0.5pt to 200pt. A band starting exactly
    /// at the line's bottom leaves it alone.
    #[test]
    fn a_full_width_band_reaching_a_paragraph_moves_its_first_line_below() {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let measure = |offset: f64, height: f64, before: f64| {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind":"shape","id":"band","shapeType":"rect","geometryPath":[],"children":[],
                 "width":200,"height":height,"wrapType":"topAndBottom",
                 "position":{"horizontal":{"relativeTo":"column","posOffset":0},
                             "vertical":{"relativeTo":"paragraph","posOffset":offset}}},
                {"kind":"paragraph","id":"anchor","attrs":{"spacing":{"before":before}},
                 "runs":[{"kind":"text","text":"anchor"}]},
                {"kind":"paragraph","id":"tail","attrs":{"spacing":{"before":before}},
                 "runs":[{"kind":"text","text":"tail"}]}
            ]))
            .unwrap();
            let measures =
                measure_blocks_with_floats(&mut blocks, &[200.0; 3], &config, None).unwrap();
            let skips: Vec<f64> = measures[1..]
                .iter()
                .map(|measure| {
                    let BlockExtent::Paragraph(paragraph) = measure else {
                        panic!()
                    };
                    paragraph.lines[0].float_skip_before.unwrap_or(0.0)
                })
                .collect();
            (skips[0], skips[1])
        };
        let close = |actual: f64, expected: f64, label: &str| {
            assert!(
                (actual - expected).abs() < 1e-3,
                "{label}: {actual} vs {expected}"
            );
        };
        let before = 4.5;
        let line = {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind":"paragraph","id":"anchor","runs":[{"kind":"text","text":"anchor"}]}
            ]))
            .unwrap();
            let BlockExtent::Paragraph(paragraph) =
                &measure_blocks_with_floats(&mut blocks, &[200.0], &config, None).unwrap()[0]
            else {
                panic!()
            };
            paragraph.lines[0].line_height
        };
        let twip = 0.05 * 96.0 / 72.0;

        // A band starting at the line's bottom clears it; one twip higher does not.
        close(measure(before + line, 1.0, before).0, 0.0, "at the bottom");
        let overlapping = before + line - twip;
        close(
            measure(overlapping, 1.0, before).0,
            overlapping + 1.0,
            "one twip into the line",
        );
        // The push clears the whole band, whatever its height.
        for height in [0.5, 8.0, 96.0, 266.0] {
            close(
                measure(before, height, before).0,
                before + height,
                "band height",
            );
        }
        // Space-before alone reaching a band moves the line that follows it.
        let (anchor_skip, tail_skip) = measure(1.0, 1.0, before);
        close(anchor_skip, 2.0, "space-before overlap");
        close(tail_skip, 0.0, "tail clear of the band");
        // A band clear of the paragraph top and of the line leaves both alone.
        let above = measure(-4.0, 1.0, before);
        close(above.0, 0.0, "band above the paragraph");
        close(above.1, 0.0, "tail below a band above the paragraph");

        // A band reaching only the second line moves that line to the band
        // bottom, with no space-before spent under it.
        let wrapped = |offset: f64| {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
                {"kind":"shape","id":"band","shapeType":"rect","geometryPath":[],"children":[],
                 "width":200,"height":2,"wrapType":"topAndBottom",
                 "position":{"horizontal":{"relativeTo":"column","posOffset":0},
                             "vertical":{"relativeTo":"paragraph","posOffset":offset}}},
                {"kind":"paragraph","id":"anchor","attrs":{"spacing":{"before":before}},
                 "runs":[{"kind":"text","text":"alpha beta gamma delta epsilon zeta eta theta"}]}
            ]))
            .unwrap();
            let BlockExtent::Paragraph(paragraph) =
                &measure_blocks_with_floats(&mut blocks, &[200.0; 2], &config, None).unwrap()[1]
            else {
                panic!()
            };
            (
                paragraph.lines[0].float_skip_before.unwrap_or(0.0),
                paragraph.lines[1].float_skip_before.unwrap_or(0.0),
            )
        };
        let second_bottom = before + 2.0 * line;
        close(wrapped(second_bottom).1, 0.0, "at the second line's bottom");
        let reaching = second_bottom - twip;
        close(
            wrapped(reaching).1,
            reaching + 2.0 - before - line,
            "one twip into the second line",
        );
    }

    /// Word 16.113 paints a `wrapNone` anchor over its cell and leaves the row
    /// alone: probing a two-cell table with and without the anchor kept the row
    /// at the same bottom and the flow below it unmoved, while the same probe
    /// with `wrapSquare` grew the row to the float's bottom.
    #[test]
    fn a_wrap_none_cell_anchor_leaves_the_row_height_alone() {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let cell_blocks = |anchor: Option<Value>| {
            let mut blocks = Vec::new();
            if let Some(anchor) = anchor {
                blocks.push(anchor);
            }
            blocks.push(
                json!({"kind":"paragraph","id":"body","runs":[{"kind":"text","text":"cell"}]}),
            );
            json!({"kind":"table","id":"outer","columnWidths":[220],
                   "rows":[{"id":"row","cells":[{"id":"cell","blocks":blocks}]}]})
        };
        let height = |value: Value| {
            let mut block: LayoutBlock = serde_json::from_value(value).unwrap();
            let measure = measure_block(&mut block, 220.0, &config).unwrap();
            let (LayoutBlock::Table(table), BlockExtent::Table(extent)) = (&block, &measure) else {
                panic!()
            };
            let layout = crate::cell_layout::layout_cell_content(
                Some(&table.rows[0].cells[0].blocks),
                Some(&extent.rows[0].cells[0].blocks),
                0.0,
            );
            assert!((layout.content_height - extent.total_height).abs() < 0.01);
            extent.total_height
        };
        let plain = height(cell_blocks(None));
        let anchor = |wrap: &str| {
            json!({"kind":"shape","id":"ellipse","shapeType":"ellipse","geometryPath":[],
                   "children":[],"width":80,"height":40,"wrapType":wrap,
                   "position":{"horizontal":{"relativeTo":"column","posOffset":0},
                               "vertical":{"relativeTo":"paragraph","posOffset":36}}})
        };
        assert_eq!(height(cell_blocks(Some(anchor("inFront")))), plain);
        assert_eq!(height(cell_blocks(Some(anchor("behind")))), plain);
        assert_eq!(height(cell_blocks(Some(anchor("square")))), plain + 40.0);
    }

    /// A wrap-only anchor must not slide the frame out from under a band
    /// anchored beside it.
    #[test]
    fn a_wrap_only_anchor_does_not_advance_the_anchor_frame() {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {"kind":"shape","id":"band","shapeType":"rect","geometryPath":[],"children":[],
             "width":40,"height":40,"wrapType":"topAndBottom",
             "position":{"horizontal":{"relativeTo":"column","posOffset":0},"vertical":{"relativeTo":"paragraph","posOffset":0}}},
            {"kind":"shape","id":"wrapOnly","shapeType":"rect","geometryPath":[],"children":[],
             "width":40,"height":40,"wrapType":"topAndBottom"},
            {"kind":"paragraph","id":"body","runs":[{"kind":"text","text":"words words words"}]}
        ]))
        .unwrap();
        let measures = measure_blocks_with_floats(&mut blocks, &[200.0; 3], &config, None).unwrap();
        let BlockExtent::Paragraph(paragraph) = &measures[2] else {
            panic!()
        };
        assert_eq!(paragraph.lines[0].float_skip_before.unwrap_or(0.0), 40.0);
    }

    #[test]
    fn vertical_table_labels_keep_a_single_rotated_line() {
        for (direction, rotation) in [("btLr", -90.0), ("tbRl", 90.0)] {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
                "kind":"table","id":"table","columnWidths":[30],"rows":[{
                    "id":"row","height":150,"heightRule":"exact","cells":[{
                        "id":"cell","textDirection":direction,"padding":{"left":0,"right":0,"top":0,"bottom":0},
                        "blocks":[{"kind":"paragraph","id":"label","runs":[{"kind":"text","text":"Vertical label","fontSize":12}]}]
                    }]
                }]
            }])).unwrap();
            let measured =
                measure_blocks(&mut blocks, 200.0, &MeasurementConfig::default()).unwrap();
            let BlockExtent::Table(table) = &measured[0] else {
                panic!()
            };
            let BlockExtent::Paragraph(label) = &table.rows[0].cells[0].blocks[0] else {
                panic!()
            };
            assert_eq!(label.lines.len(), 1);
            assert_eq!(table.total_height, 150.0);
            let mut input = crate::types::Input {
                measured: blocks
                    .into_iter()
                    .zip(measured)
                    .map(|(block, measure)| crate::types::MeasuredBlock { block, measure })
                    .collect(),
                options: crate::types::LayoutOptions::default(),
            };
            let layout = crate::compute_layout_input(&mut input).unwrap();
            let display = crate::build_display_list(&input, &layout).unwrap();
            let text = display.pages[0]
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    crate::display_list::Primitive::Text(text) if text.text == "Vertical label" => {
                        Some(text)
                    }
                    _ => None,
                })
                .unwrap();
            assert_eq!(
                text.rotation_deg
                    .as_ref()
                    .and_then(serde_json::Number::as_f64),
                Some(rotation)
            );
        }
    }

    #[test]
    fn vertical_labels_wrap_at_the_height_of_the_rows_they_span() {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font_id])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..MeasurementConfig::default()
        };
        let pad = json!({"left":0,"right":0,"top":4,"bottom":4});
        let label = |text: &str, span: u32| {
            json!({"id":"label","textDirection":"btLr","rowSpan":span,"padding":pad,
            "blocks":[{"kind":"paragraph","id":"l","runs":[{"kind":"text","text":text,"fontSize":12}]}]})
        };
        let text = |id: &str, lines: usize| {
            json!({"id":id,"padding":pad,"blocks":(0..lines)
            .map(|line| json!({"kind":"paragraph","id":format!("{id}{line}"),"runs":[{"kind":"text","text":"Row","fontSize":12}]}))
            .collect::<Vec<_>>()})
        };
        let measure = |rows: serde_json::Value| {
            let mut blocks: Vec<LayoutBlock> = serde_json::from_value(
                json!([{"kind":"table","id":"table","columnWidths":[30, 100],"rows":rows}]),
            )
            .unwrap();
            let BlockExtent::Table(table) = measure_blocks(&mut blocks, 200.0, &config)
                .unwrap()
                .remove(0)
            else {
                panic!()
            };
            table
        };
        let lines = |table: &TableExtent| {
            let BlockExtent::Paragraph(label) = &table.rows[0].cells[0].blocks[0] else {
                panic!()
            };
            label.lines.clone()
        };

        let single = measure(json!([{"id":"r0","height":40,"heightRule":"atLeast",
            "cells":[label("Vertical label", 1), text("a", 1)]}]));
        let longest = lines(&single)
            .iter()
            .map(|line| line.width)
            .fold(0.0, f64::max);
        assert!(lines(&single).len() > 1);
        assert_eq!(single.rows[0].height, (longest + 8.0).max(48.0));

        let spanned = measure(json!([
            {"id":"r0","height":20,"heightRule":"atLeast","cells":[label("Vertical label", 2), text("a", 5)]},
            {"id":"r1","height":20,"heightRule":"atLeast","cells":[text("b", 1)]}
        ]));
        let rows_height = spanned.rows[0].height + spanned.rows[1].height;
        assert_eq!(lines(&spanned).len(), 1);
        assert!(lines(&spanned)[0].width <= rows_height - 8.0);
        assert_eq!(spanned.rows[1].height, 28.0);

        let tall = measure(json!([{"id":"r0","height":500,"heightRule":"atLeast",
            "cells":[label("A vertical label longer than the table is wide", 1), text("a", 1)]}]));
        assert_eq!(lines(&tall).len(), 1);
        assert!(lines(&tall)[0].width > 200.0);
        assert_eq!(tall.rows[0].height, 508.0);
    }

    #[test]
    fn percentage_width_nested_keeps_compat_basis_in_unshifted_story() {
        for measure in [measure_blocks, measure_blocks_without_table_compat_shift] {
            for algorithm in [None, Some("autofit"), Some("fixed")] {
                let nested = json!({
                    "kind": "table", "id": "nested", "compatibilityMode": 14,
                    "cellMarginLeft": 7.2, "cellMarginRight": 7.2,
                    "width": 5000, "widthType": "pct", "widthAlgorithm": algorithm,
                    "columnWidths": [100, 100], "rows": [{"id": "nested-row", "cells": [
                        {"id": "left", "blocks": [], "minContentWidth": 20, "maxContentWidth": 400,
                         "padding": {"top": 0, "bottom": 0, "left": 7.2, "right": 0}},
                        {"id": "right", "blocks": [], "minContentWidth": 20, "maxContentWidth": 400,
                         "padding": {"top": 0, "bottom": 0, "left": 0, "right": 7.2}}
                    ]}]
                });
                let mut blocks = serde_json::from_value::<Vec<LayoutBlock>>(json!([{
                    "kind": "table", "id": "outer", "width": 8313, "widthType": "dxa",
                    "columnWidths": [554.2], "rows": [{"id": "outer-row", "cells": [{
                        "id": "outer-cell", "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
                        "blocks": [nested]
                    }]}]
                }]))
                .unwrap();
                let measured = measure(&mut blocks, 554.2, &MeasurementConfig::default()).unwrap();
                let BlockExtent::Table(outer) = &measured[0] else {
                    panic!("outer table expected");
                };
                let BlockExtent::Table(nested) = &outer.rows[0].cells[0].blocks[0] else {
                    panic!("nested table expected");
                };
                assert!((outer.total_width - 554.2).abs() < 1e-6);
                assert!(
                    (nested.total_width - 568.6).abs() < 1e-6,
                    "{algorithm:?}: nested width {}, expected 568.6",
                    nested.total_width
                );
            }
        }
    }

    #[test]
    fn collapsed_borders_expand_minimum_and_auto_rows_but_not_exact_rows() {
        for (height, rule, expected) in [
            (Some(40.0), "atLeast", 42.0),
            (Some(40.0), "exact", 41.0),
            (None, "auto", 18.0),
        ] {
            let mut table: TableBlock = serde_json::from_value(json!({
                "kind":"table", "id":"table", "columnWidths":[100],
                "rows":[{"id":"row", "height":height, "heightRule":rule, "cells":[{
                    "id":"cell", "padding":{"top":0,"bottom":0,"left":0,"right":0},
                    "borders":{"top":{"width":1},"bottom":{"width":1}},
                    "blocks":[{"kind":"image","id":"image","src":"","width":10,"height":16}]
                }]}]
            }))
            .unwrap();
            let measured = measure_table(&mut table, 100.0, &MeasurementConfig::default()).unwrap();
            assert_eq!(measured.rows[0].height, expected);
            assert_eq!(measured.total_height, expected);
        }
    }

    #[test]
    fn outer_borders_expand_only_the_real_non_exact_outer_rows() {
        for (exact_rows, expected) in [
            ([false, false, false], [19.0, 18.0, 19.0]),
            ([true, false, false], [40.0, 18.0, 19.0]),
            ([false, true, false], [19.0, 40.0, 19.0]),
            ([false, false, true], [19.0, 18.0, 42.0]),
        ] {
            let rows: Vec<_> = exact_rows
                .into_iter()
                .enumerate()
                .map(|(index, exact)| {
                    json!({
                        "id": index, "height": if exact { Some(40) } else { None },
                        "heightRule": if exact { "exact" } else { "auto" },
                        "cells": [{
                            "id": index + 3,
                            "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
                            "borders": {"top": {"width": 2}, "bottom": {"width": 2}},
                            "blocks": [{"kind": "image", "id": index + 6,
                                "src": "", "width": 10, "height": 16}]
                        }]
                    })
                })
                .collect();
            let mut table: TableBlock = serde_json::from_value(json!({
                "id": "table", "columnWidths": [100], "rows": rows
            }))
            .unwrap();
            let measured = measure_table(&mut table, 100.0, &MeasurementConfig::default()).unwrap();
            let heights: Vec<_> = measured.rows.iter().map(|row| row.height).collect();
            assert_eq!(heights, expected);
            assert_eq!(measured.total_height, expected.iter().sum::<f64>());
            let breaks = crate::table_row_break::build_table_row_break_info(&table, &measured);
            assert_eq!(
                breaks.row_tops,
                vec![
                    0.0,
                    expected[0],
                    expected[0] + expected[1],
                    measured.total_height
                ]
            );
            assert_eq!(
                crate::table_row_break::first_table_fragment_height(&table, &measured, &breaks),
                expected[0]
            );
        }
    }

    #[test]
    fn invisible_outer_borders_reserve_no_band() {
        let mut table: TableBlock = serde_json::from_value(json!({
            "id":"table", "columnWidths":[100],
            "rows":[{"id":"row", "height":40, "heightRule":"exact", "cells":[{
                "id":"cell", "padding":{"top":0,"bottom":0,"left":0,"right":0},
                "borders":{"top":{"width":6,"style":"nil"},"bottom":{"width":6,"style":"none"}},
                "blocks":[{"kind":"image","id":"image","src":"","width":10,"height":16}]
            }]}]
        }))
        .unwrap();
        let measured = measure_table(&mut table, 100.0, &MeasurementConfig::default()).unwrap();
        assert_eq!(measured.total_height, 40.0);
    }

    #[test]
    fn merged_cells_share_row_growth_without_changing_exact_rows() {
        for (heights, exact, expected) in [
            ([128.0, 160.0], [false, false, false], [16.0, 16.0, 128.0]),
            ([160.0, 128.0], [false, false, false], [16.0, 16.0, 128.0]),
            ([160.0, 160.0], [false, false, false], [16.0, 16.0, 128.0]),
            ([16.0, 32.0], [false, false, false], [16.0, 16.0, 16.0]),
            ([128.0, 160.0], [false, false, true], [16.0, 128.0, 16.0]),
            ([160.0, 128.0], [true, false, true], [16.0, 128.0, 16.0]),
            ([128.0, 160.0], [false, true, true], [128.0, 16.0, 16.0]),
            ([128.0, 160.0], [true, true, true], [16.0, 16.0, 16.0]),
        ] {
            let cell = |id: &str, height: f64, row_span: usize| {
                json!({
                    "id":id, "rowSpan":row_span,
                    "padding":{"top":0,"bottom":0,"left":0,"right":0},
                    "blocks":[{"kind":"image","id":id,"src":"","width":10,"height":height}]
                })
            };
            let rows: Vec<_> = exact
                .iter()
                .enumerate()
                .map(|(index, exact)| {
                    let mut cells = Vec::new();
                    if index == 0 {
                        cells.push(cell("left", heights[0], 3));
                        cells.push(cell("right", heights[1], 3));
                    }
                    cells.push(cell(&format!("marker{index}"), 16.0, 1));
                    json!({
                        "id":format!("row{index}"), "height":16,
                        "heightRule":if *exact { "exact" } else { "atLeast" }, "cells":cells
                    })
                })
                .collect();
            let mut table: TableBlock = serde_json::from_value(json!({
                "id":"table", "columnWidths":[100,100,100], "rows":rows
            }))
            .unwrap();
            let measured = measure_table(&mut table, 300.0, &MeasurementConfig::default()).unwrap();
            let actual: Vec<_> = measured.rows.iter().map(|row| row.height).collect();
            assert_eq!(actual, expected, "heights {heights:?}, exact {exact:?}");
            assert_eq!(measured.total_height, expected.iter().sum::<f64>());
        }
    }

    #[test]
    fn overlapping_merged_cells_use_growth_in_their_shared_rows() {
        let cell = |id: &str, height: f64, row_span: usize, column: usize| {
            json!({
                "id":id, "rowSpan":row_span, "gridStart":column,
                "padding":{"top":0,"bottom":0,"left":0,"right":0},
                "blocks":[{"kind":"image","id":id,"src":"","width":10,"height":height}]
            })
        };
        let mut table: TableBlock = serde_json::from_value(json!({
            "id":"table", "columnWidths":[100,100,100], "rows":[
                {"id":"row0", "cells":[cell("left", 80.0, 3, 0), cell("marker0", 16.0, 1, 2)]},
                {"id":"row1", "cells":[cell("right", 96.0, 3, 1), cell("marker1", 16.0, 1, 2)]},
                {"id":"row2", "cells":[cell("marker2", 16.0, 1, 2)]},
                {"id":"row3", "cells":[cell("marker3", 16.0, 1, 2)]}
            ]
        }))
        .unwrap();
        let measured = measure_table(&mut table, 300.0, &MeasurementConfig::default()).unwrap();
        let actual: Vec<_> = measured.rows.iter().map(|row| row.height).collect();
        assert_eq!(actual, [16.0, 16.0, 48.0, 32.0]);
        assert_eq!(measured.total_height, 112.0);
    }

    #[test]
    fn unequal_merged_spans_share_growth_in_either_column_order() {
        let cell = |id: &str, height: f64, row_span: usize, column: usize| {
            json!({
                "id":id, "rowSpan":row_span, "gridStart":column,
                "padding":{"top":0,"bottom":0,"left":0,"right":0},
                "blocks":[{"kind":"image","id":id,"src":"","width":10,"height":height}]
            })
        };
        for spans in [[(96.0, 3), (80.0, 2)], [(80.0, 2), (96.0, 3)]] {
            let mut table: TableBlock = serde_json::from_value(json!({
                "id":"table", "columnWidths":[100,100,100], "rows":[
                    {"id":"row0", "cantSplit":true, "cells":[
                        cell("left", spans[0].0, spans[0].1, 0),
                        cell("right", spans[1].0, spans[1].1, 1),
                        cell("marker0", 16.0, 1, 2)
                    ]},
                    {"id":"row1", "cantSplit":true, "cells":[cell("marker1", 16.0, 1, 2)]},
                    {"id":"row2", "cantSplit":true, "cells":[cell("marker2", 16.0, 1, 2)]}
                ]
            }))
            .unwrap();
            let measured = measure_table(&mut table, 300.0, &MeasurementConfig::default()).unwrap();
            let actual: Vec<_> = measured.rows.iter().map(|row| row.height).collect();
            assert_eq!(actual, [16.0, 64.0, 16.0], "spans {spans:?}");
            assert_eq!(measured.total_height, 96.0);
            let mut input = crate::types::Input {
                measured: vec![crate::types::MeasuredBlock {
                    block: LayoutBlock::Table(table),
                    measure: BlockExtent::Table(measured),
                }],
                options: serde_json::from_value(json!({"pageSize":{"w":300,"h":120},
                    "margins":{"top":0,"bottom":0,"left":0,"right":0}}))
                .unwrap(),
            };
            let layout = crate::compute_layout_input(&mut input).unwrap();
            assert_eq!(layout.pages.len(), 1, "spans {spans:?}");
        }
    }

    #[test]
    fn minimum_row_height_reserves_cell_margins_outside_the_content_minimum() {
        for (minimum, content, expected) in [(40, 16, 57.0), (80, 16, 97.0), (40, 60, 77.0)] {
            let mut table: TableBlock = serde_json::from_value(json!({
                "id":"table", "columnWidths":[100],
                "rows":[{"id":"row", "height":minimum, "heightRule":"atLeast", "cells":[{
                    "id":"cell", "padding":{"top":5,"bottom":10,"left":0,"right":0},
                    "borders":{"top":{"width":1},"bottom":{"width":1}},
                    "blocks":[{"kind":"image","id":"image","src":"","width":10,"height":content}]
                }]}]
            }))
            .unwrap();
            let measured = measure_table(&mut table, 100.0, &MeasurementConfig::default()).unwrap();
            assert_eq!(measured.rows[0].height, expected);
            assert_eq!(measured.rows[0].cells[0].height, content as f64 + 15.0);
        }
    }

    #[test]
    fn rotated_image_only_cells_use_the_visual_height() {
        for (rotation, expected) in [(None, 102.0), (Some(90), 52.0), (Some(270), 52.0)] {
            let bounds = rotation.map(|_| json!({"width":80,"height":30}));
            let mut table: TableBlock = serde_json::from_value(json!({
                "id":"table", "columnWidths":[100],
                "rows":[{"id":"row", "cells":[{
                    "id":"cell", "padding":{"top":5,"bottom":7,"left":0,"right":0},
                    "blocks":[{"kind":"paragraph","id":"paragraph",
                        "attrs":{"spacing":{"before":4,"after":6}},
                        "runs":[{"kind":"image","src":"","width":30,"height":80,
                            "rotationDeg":rotation,"rotationBounds":bounds}]}]
                }]}]
            }))
            .unwrap();
            let measured = measure_table(&mut table, 100.0, &MeasurementConfig::default()).unwrap();
            assert_eq!(measured.rows[0].height, expected);
            assert_eq!(measured.rows[0].cells[0].height, expected);
        }
    }

    #[test]
    fn rotated_images_mixed_with_text_keep_the_measured_line_height() {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind":"paragraph","id":"paragraph","runs":[
                {"kind":"text","text":"label"},
                {"kind":"image","src":"","width":30,"height":80,
                    "rotationDeg":270,"rotationBounds":{"width":80,"height":30}}
            ]
        }))
        .unwrap();
        let measure: BlockExtent = serde_json::from_value(json!({
            "kind":"paragraph","totalHeight":44,"lines":[{
                "headRun":0,"headChar":0,"tailRun":1,"tailChar":1,
                "width":90,"ascent":30,"descent":4,"lineHeight":44
            }]
        }))
        .unwrap();
        assert_eq!(table_cell_block_height(&block, &measure), 44.0);
    }

    #[test]
    fn measures_non_text_blocks_without_host_callbacks() {
        let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {"kind": "image", "id": "i", "src": "x", "width": 10, "height": 20},
            {"kind": "chart", "id": "c", "chart": {}, "width": 30, "height": 40},
            {"kind": "pageBreak", "id": "b"}
        ]))
        .unwrap();

        let measured = measure_blocks(&mut blocks, 100.0, &MeasurementConfig::default()).unwrap();
        assert!(matches!(
            measured[0],
            BlockExtent::Image(ImageExtent {
                width: 10.0,
                height: 20.0
            })
        ));
        assert!(matches!(
            measured[1],
            BlockExtent::Chart(ChartExtent {
                width: 30.0,
                height: 40.0
            })
        ));
        assert!(matches!(measured[2], BlockExtent::PageBreak));
    }

    #[test]
    fn collects_nested_font_styles_and_script_fallbacks() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
            "kind": "table",
            "id": "t",
            "rows": [{
                "id": "r",
                "cells": [{
                    "id": "c",
                    "blocks": [{
                        "kind": "paragraph",
                        "id": "p",
                        "runs": [{
                            "kind": "text",
                            "text": "Latin 日本語かな",
                            "fontFamily": "Aptos",
                            "bold": true,
                            "fontSlots": {"eastAsia": "Yu Mincho"}
                        }]
                    }]
                }]
            }]
        }]))
        .unwrap();

        let requirements = collect_font_requirements(&blocks, "Calibri");
        let value = serde_json::to_value(requirements).unwrap();

        assert!(value.as_array().unwrap().iter().any(|requirement| {
            requirement["key"] == "aptos|1|0"
                && requirement["scripts"] == serde_json::json!(["cjk-jp"])
        }));
        assert!(value.as_array().unwrap().iter().any(|requirement| {
            requirement["key"] == "yu mincho|0|0"
                && requirement["scripts"] == serde_json::json!(["cjk-jp"])
        }));
    }

    fn font_requirement_keys(run: Value) -> Vec<String> {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph", "id": "p", "runs": [run]
        }))
        .unwrap();
        collect_font_requirements([&block], "Calibri")
            .into_iter()
            .map(|requirement| requirement.key)
            .collect()
    }

    #[test]
    fn collects_list_marker_face_from_the_first_unnamed_text_run() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {
                "kind": "paragraph", "id": "list", "attrs": {
                    "listMarker": "1.", "listMarkerBold": true, "indent": {"hanging": 0}
                },
                "runs": [
                    {"kind": "text", "text": "First"},
                    {"kind": "text", "text": "Later", "fontFamily": "Aptos"}
                ]
            },
            {
                "kind": "paragraph", "id": "unused", "runs": [{
                    "kind": "text", "text": "Latin", "fontFamily": "Arial",
                    "boldCs": true, "fontSlots": {"cs": "Calibri"}
                }]
            }
        ]))
        .unwrap();
        let requirements = collect_font_requirements(&blocks, "Calibri");

        assert!(
            requirements
                .iter()
                .any(|requirement| requirement.key == "calibri|1|0")
        );
        let mut named = BTreeMap::new();
        collect_font_requirements_into(&blocks, "Calibri", &mut named);
        assert!(named.contains_key("aptos|1|0"));
    }

    #[test]
    fn preserves_script_order_from_unused_font_slots() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {
                "kind": "paragraph", "id": "jp", "runs": [{
                    "kind": "text", "text": "かな", "fontFamily": "Aptos",
                    "boldCs": true, "fontSlots": {"cs": "Arial"}
                }]
            },
            {
                "kind": "paragraph", "id": "sc", "runs": [{
                    "kind": "text", "text": "漢字", "fontFamily": "Arial"
                }]
            }
        ]))
        .unwrap();
        let mut split = FontRequirementCollector::default();
        for block in &blocks {
            split.collect([block], "Calibri");
        }
        for requirements in [
            collect_font_requirements(&blocks, "Calibri"),
            split.finish().into_values().collect(),
        ] {
            let keys: Vec<_> = requirements
                .iter()
                .map(|requirement| requirement.key.as_str())
                .collect();

            assert_eq!(keys, ["aptos|0|0", "arial|0|0", "calibri|0|0"]);
            assert_eq!(requirements[0].scripts, ["cjk-jp"]);
            assert_eq!(requirements[1].scripts, ["cjk-jp", "cjk-sc"]);
            assert_eq!(requirements[2].scripts, ["cjk-jp", "cjk-sc"]);
        }
    }

    #[test]
    fn keeps_named_script_order_for_a_face_an_unnamed_slot_reaches() {
        let blocks: Vec<LayoutBlock> = [
            ("かな", json!({"hAnsi": "Arial"})),
            ("漢字", json!({"cs": "Arial"})),
            ("かな", json!({"cs": "Arial"})),
        ]
        .into_iter()
        .enumerate()
        .map(|(index, (text, slots))| {
            serde_json::from_value(json!({
                "kind": "paragraph", "id": format!("p{index}"), "runs": [{
                    "kind": "text", "text": text, "complexScript": true,
                    "boldCs": true, "fontSlots": slots
                }]
            }))
            .unwrap()
        })
        .collect();
        let requirements = collect_font_requirements(&blocks, "Calibri");
        let arial_bold = requirements
            .iter()
            .find(|requirement| requirement.key == "arial|1|0")
            .unwrap();
        let mut split = BTreeMap::new();
        for block in &blocks {
            collect_font_requirements_into([block], "Calibri", &mut split);
        }

        assert_eq!(arial_bold.scripts, ["cjk-sc", "cjk-jp"]);
        assert_eq!(split["arial|1|0"].scripts, ["cjk-sc", "cjk-jp"]);
    }

    #[test]
    fn collects_east_asian_font_slots_only_for_east_asian_text() {
        for (text, expected) in [
            (
                "Latin",
                vec![
                    "aptos|0|0",
                    "arial|0|0",
                    "calibri|0|0",
                    "times new roman|0|0",
                ],
            ),
            (
                "漢字",
                vec![
                    "aptos|0|0",
                    "arial|0|0",
                    "calibri|0|0",
                    "simsun|0|0",
                    "times new roman|0|0",
                ],
            ),
        ] {
            let keys = font_requirement_keys(json!({
                "kind": "text", "text": text, "fontFamily": "Aptos",
                "fontSlots": {
                    "ascii": "Arial", "hAnsi": "Times New Roman",
                    "eastAsia": "SimSun", "cs": "Traditional Arabic"
                }
            }));
            assert_eq!(keys, expected, "{text}");
        }
    }

    #[test]
    fn collects_east_asian_font_slots_for_hinted_text() {
        for (hint, expected) in [
            (None, vec!["calibri|0|0"]),
            (Some("eastAsia"), vec!["calibri|0|0", "simsun|0|0"]),
        ] {
            let keys = font_requirement_keys(json!({
                "kind": "text", "text": "Latin \u{201c}",
                "fontSlots": {"eastAsia": "SimSun", "hint": hint}
            }));
            assert_eq!(keys, expected, "{hint:?}");
        }
    }

    #[test]
    fn collects_complex_script_font_slots_only_when_used() {
        for (complex_script, expected) in [
            (None, vec!["calibri|0|0", "calibri|1|0"]),
            (Some(false), vec!["calibri|0|0", "calibri|1|0"]),
            (
                Some(true),
                vec![
                    "calibri|0|0",
                    "calibri|1|0",
                    "traditional arabic|0|0",
                    "traditional arabic|0|1",
                ],
            ),
        ] {
            let keys = font_requirement_keys(json!({
                "kind": "text", "text": "Latin", "bold": true,
                "boldCs": false, "italicCs": true, "complexScript": complex_script,
                "fontSlots": {"eastAsia": "SimSun", "cs": "Traditional Arabic"}
            }));
            assert_eq!(keys, expected, "{complex_script:?}");
        }
    }

    #[test]
    fn collects_complex_script_face_without_a_cs_font_slot() {
        for (slots, default_family) in [
            (json!({"ascii": "Arial", "hAnsi": "Arial"}), "Calibri"),
            (Value::Null, "Arial"),
        ] {
            let block: LayoutBlock = serde_json::from_value(json!({
                "kind": "paragraph", "id": "p", "runs": [
                    {
                        "kind": "text", "text": "Latin", "boldCs": true,
                        "fontSlots": {"cs": "Arial", "ascii": "Arial", "hAnsi": "Arial"}
                    },
                    {
                        "kind": "text", "text": "Latin", "complexScript": true, "boldCs": true,
                        "fontSlots": slots
                    }
                ]
            }))
            .unwrap();
            let keys: Vec<_> = collect_font_requirements([&block], default_family)
                .into_iter()
                .map(|requirement| requirement.key)
                .filter(|key| key.starts_with("arial"))
                .collect();

            assert_eq!(keys, ["arial|0|0", "arial|1|0"], "{default_family}");
        }
    }

    #[test]
    fn collects_document_default_for_an_unnamed_text_slot() {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph", "id": "p", "attrs": {"defaultFontFamily": "Arial"},
            "runs": [{
                "kind": "text", "text": "Latin", "fontSlots": {"cs": "Calibri"}
            }]
        }))
        .unwrap();
        let keys: Vec<_> = collect_font_requirements([&block], "Calibri")
            .into_iter()
            .map(|requirement| requirement.key)
            .collect();

        assert_eq!(keys, ["arial|0|0", "calibri|0|0"]);
    }

    #[test]
    fn tab_font_requirements_skip_script_slots() {
        for complex_script in [false, true] {
            let keys = font_requirement_keys(json!({
                "kind": "tab", "fontFamily": "Aptos", "bold": true,
                "complexScript": complex_script,
                "fontSlots": {
                    "ascii": "Arial", "hAnsi": "Times New Roman",
                    "eastAsia": "SimSun", "cs": "Traditional Arabic", "hint": "eastAsia"
                }
            }));
            let mut expected = vec![
                "aptos|1|0",
                "arial|1|0",
                "calibri|0|0",
                "times new roman|1|0",
            ];
            // Text typed after the tab inherits the run and measures with its cs slot.
            if complex_script {
                expected.push("traditional arabic|1|0");
            }
            assert_eq!(keys, expected, "{complex_script}");
        }
    }

    #[test]
    fn complex_script_tab_reaches_the_text_face_for_an_unnamed_cs_slot() {
        let keys = font_requirement_keys(json!({
            "kind": "tab", "boldCs": true, "complexScript": true,
            "fontSlots": {"ascii": "Arial", "hAnsi": "Courier New"}
        }));
        assert!(keys.contains(&"courier new|1|0".to_owned()), "{keys:?}");
    }

    #[test]
    fn collects_the_run_family_of_an_empty_text_run() {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph", "id": "p", "attrs": {"defaultFontFamily": "Arial"},
            "runs": [
                {
                    "kind": "text", "text": "", "bold": true,
                    "fontSlots": {"hAnsi": "Aptos", "cs": "Calibri"}
                },
                {"kind": "text", "text": "Latin", "fontFamily": "Arial"}
            ]
        }))
        .unwrap();
        let keys: Vec<_> = collect_font_requirements([&block], "Calibri")
            .into_iter()
            .map(|requirement| requirement.key)
            .collect();

        assert!(keys.contains(&"calibri|1|0".to_owned()), "{keys:?}");
    }

    #[test]
    fn collecting_into_a_map_keeps_scripts_of_an_existing_requirement() {
        let paragraph = |id: &str, text: &str, mut run: Value| -> LayoutBlock {
            run["kind"] = json!("text");
            run["text"] = json!(text);
            serde_json::from_value(json!({"kind": "paragraph", "id": id, "runs": [run]})).unwrap()
        };
        let han = paragraph("sc", "漢字", json!({"fontFamily": "Arial"}));
        let kana = paragraph(
            "jp",
            "かな",
            json!({"fontFamily": "Aptos", "fontSlots": {"cs": "Arial"}}),
        );
        for (order, expected) in [
            ([&han, &kana], ["cjk-sc", "cjk-jp"]),
            ([&kana, &han], ["cjk-jp", "cjk-sc"]),
        ] {
            let mut requirements = BTreeMap::new();
            for block in order {
                collect_font_requirements_into([block], "Calibri", &mut requirements);
            }
            assert_eq!(requirements["arial|0|0"].scripts, expected);
        }
    }

    #[test]
    fn collects_document_default_for_tab_runs_with_font_slots() {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph", "id": "p", "attrs": {"defaultFontFamily": "Arial"},
            "runs": [{"kind": "tab", "fontSlots": {"cs": "Calibri"}}]
        }))
        .unwrap();
        let keys: Vec<_> = collect_font_requirements([&block], "Calibri")
            .into_iter()
            .map(|requirement| requirement.key)
            .collect();

        assert_eq!(keys, ["arial|0|0", "calibri|0|0"]);
    }

    #[test]
    fn field_font_requirements_skip_script_slots() {
        for fallback in [
            None,
            Some(""),
            Some("Latin"),
            Some("漢字"),
            Some("\u{201c}"),
        ] {
            let keys = font_requirement_keys(json!({
                "kind": "field", "fieldType": "PAGE", "fallback": fallback,
                "fontFamily": "Aptos", "italic": true, "complexScript": true,
                "fontSlots": {"eastAsia": "SimSun", "cs": "Traditional Arabic", "hint": "eastAsia"}
            }));
            // A complex-script run keeps its cs slot for text typed into it.
            assert_eq!(
                keys,
                ["aptos|0|1", "calibri|0|0", "traditional arabic|0|1"],
                "{fallback:?}"
            );
        }
    }

    #[test]
    fn missing_font_chain_uses_synthetic_extent() {
        crate::clear_measure_fonts();
        let mut block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph",
            "id": "p",
            "runs": [{"kind": "text", "text": "abcd", "fontSize": 12}],
            "attrs": {"spacing": {"before": 2, "after": 3}}
        }))
        .unwrap();

        let BlockExtent::Paragraph(extent) =
            measure_block(&mut block, 100.0, &MeasurementConfig::default()).unwrap()
        else {
            panic!("paragraph expected");
        };

        assert_eq!(extent.lines[0].width, 64.0);
        assert_eq!(extent.lines[0].ascent, 12.8);
        assert_eq!(extent.lines[0].descent, 3.2);
        assert_eq!(extent.total_height, 23.4);
    }

    fn wrapped_shape(wrap_text: Option<&str>, x: f64, width: f64) -> ShapeBlock {
        serde_json::from_value(json!({
            "id": "s",
            "shapeType": "rect",
            "geometryPath": [],
            "width": width,
            "height": 40.0,
            "children": [],
            "wrapType": "square",
            "wrapText": wrap_text,
            "wrapDistances": {"top": 0, "bottom": 0, "left": 12, "right": 12},
            "position": {
                "horizontal": {"relativeTo": "column", "posOffset": x},
                "vertical": {"relativeTo": "paragraph", "posOffset": 0}
            }
        }))
        .unwrap()
    }

    fn shape_zone(wrap_text: Option<&str>, x: f64, width: f64) -> FloatingZone {
        let mut zones = Vec::new();
        extract_shape_zone(
            &wrapped_shape(wrap_text, x, width),
            0,
            600.0,
            None,
            None,
            &mut zones,
        );
        zones.pop().expect("zone").zone
    }

    #[test]
    fn a_narrow_interior_both_sides_float_keeps_a_strip_on_each_side() {
        let zone = shape_zone(None, 300.0, 6.0);
        assert_eq!(
            zone.segments
                .iter()
                .map(|strip| (strip.left_offset, strip.available_width))
                .collect::<Vec<_>>(),
            vec![(0.0, 288.0), (318.0, 282.0)]
        );
        assert_eq!((zone.left_margin, zone.right_margin), (0.0, 0.0));
    }

    #[test]
    fn a_float_wider_than_the_side_it_would_cost_keeps_one_side() {
        let zone = shape_zone(None, 200.0, 220.0);
        assert!(zone.segments.is_empty());
        assert_eq!((zone.left_margin, zone.right_margin), (0.0, 412.0));
    }

    #[test]
    fn largest_and_one_sided_wraps_never_split_the_line() {
        for wrap_text in ["largest", "left", "right"] {
            assert!(
                shape_zone(Some(wrap_text), 300.0, 6.0).segments.is_empty(),
                "{wrap_text} must keep a single side"
            );
        }
    }
    /// A stale `w:gridCol` must not force a wrap: the heading in
    /// `oxi-en-administrative-04` fits one line in Word only because Word
    /// re-measures the `auto` column the declared grid left 5.84px short.
    #[test]
    fn a_stale_grid_column_widens_to_keep_its_heading_on_one_line() {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..Default::default()
        };
        let padding = json!({"top":0,"bottom":0,"left":1,"right":1});
        let cell = |text: &str, width: Value| {
            json!({"id":text,"padding":padding,"widthValue":width,"widthType":"dxa","blocks":[
                {"kind":"paragraph","id":text,"runs":[{"kind":"text","text":text}]}]})
        };
        let mut block: LayoutBlock = serde_json::from_value(json!({
            "kind":"table","id":"stale","columnWidths":[60,100],
            "rows":[{"id":"row","cells":[
                cell("Content sized column", json!(0)),
                cell("Priced", json!(1500)),
            ]}]
        }))
        .unwrap();
        let measure = measure_block(&mut block, 300.0, &config).unwrap();
        let BlockExtent::Table(extent) = &measure else {
            panic!()
        };
        let heading = &extent.rows[0].cells[0];
        let BlockExtent::Paragraph(paragraph) = &heading.blocks[0] else {
            panic!()
        };
        assert_eq!(paragraph.lines.len(), 1);
        assert!(heading.width > 60.0 && heading.width < 300.0);
        assert!((heading.width - (paragraph.lines[0].width + 2.0)).abs() < 0.01);
        assert_eq!(extent.rows[0].cells[1].width, 100.0);
    }

    fn word_cell(text: &str, span: usize) -> serde_json::Value {
        word_cell_paragraphs(&[text], span)
    }

    fn word_cell_paragraphs(texts: &[&str], span: usize) -> serde_json::Value {
        let blocks: Vec<_> = texts
            .iter()
            .map(|text| json!({"kind":"paragraph","id":text,"runs":[{"kind":"text","text":text}]}))
            .collect();
        json!({"id":texts.join("|"),"colSpan":span,"padding":{"top":0,"bottom":0,"left":1,"right":1},"blocks":blocks})
    }

    fn word_table_rows(
        layout: Option<&str>,
        rows: Vec<Vec<serde_json::Value>>,
    ) -> (TableExtent, ParagraphExtent) {
        word_table_row_objects(
            layout,
            rows.into_iter()
                .map(|cells| json!({"cells":cells}))
                .collect(),
        )
    }

    fn word_table_row_objects(
        layout: Option<&str>,
        rows: Vec<serde_json::Value>,
    ) -> (TableExtent, ParagraphExtent) {
        let font = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..Default::default()
        };
        let rows: Vec<_> = rows
            .into_iter()
            .enumerate()
            .map(|(index, mut row)| {
                row["id"] = json!(format!("row{index}"));
                row
            })
            .collect();
        let mut table = json!({
            "kind":"table","id":"words","columnWidths":[60,100,100],"width":3900,"widthType":"dxa",
            "rows":rows
        });
        if let Some(layout) = layout {
            table["layoutMode"] = json!(layout);
            table["widthAlgorithm"] = json!("legacy");
        }
        let mut block: LayoutBlock = serde_json::from_value(table).unwrap();
        let BlockExtent::Table(extent) = measure_block(&mut block, 300.0, &config).unwrap() else {
            panic!()
        };
        let BlockExtent::Paragraph(paragraph) = extent.rows[0].cells[0].blocks[0].clone() else {
            panic!()
        };
        (extent, paragraph)
    }

    fn word_table(layout: Option<&str>) -> (TableExtent, ParagraphExtent) {
        word_table_rows(
            layout,
            vec![vec![
                word_cell("0000000000", 1),
                word_cell("00 00", 1),
                word_cell("0 0", 1),
            ]],
        )
    }

    #[test]
    fn an_autofit_column_widens_to_its_longest_word_and_the_others_give_equal_shares() {
        let word = 10.0 * 1139.0 / 128.0 + 2.0;
        for layout in [None, Some("autofit")] {
            let (extent, paragraph) = word_table(layout);
            assert_eq!(paragraph.lines.len(), 1);
            assert!((extent.column_widths[0] - word).abs() < 1e-3);
            let share = (word - 60.0) / 2.0;
            assert!((extent.column_widths[1] - (100.0 - share)).abs() < 1e-3);
            assert!((extent.column_widths[2] - (100.0 - share)).abs() < 1e-3);
            assert!((extent.total_width - 260.0).abs() < 1e-6);
        }
    }

    #[test]
    fn a_fixed_layout_column_keeps_its_width_and_wraps_the_word() {
        let (extent, paragraph) = word_table(Some("fixed"));
        assert_eq!(extent.column_widths, vec![60.0, 100.0, 100.0]);
        assert_eq!(paragraph.lines.len(), 2);
    }

    #[test]
    fn a_no_break_space_keeps_its_word_together_when_the_column_widens() {
        let word = 9.0 * 1139.0 / 128.0 + 569.0 / 128.0 + 2.0;
        let (extent, paragraph) = word_table_rows(
            None,
            vec![vec![
                word_cell("000000\u{00A0}000", 1),
                word_cell("00 00", 1),
                word_cell("0 0", 1),
            ]],
        );
        assert_eq!(paragraph.lines.len(), 1);
        assert!((extent.column_widths[0] - word).abs() < 1e-3);
    }

    #[test]
    fn columns_keep_their_widths_when_widening_would_break_a_word_in_a_spanning_cell() {
        let (extent, paragraph) = word_table_rows(
            None,
            vec![
                vec![
                    word_cell("0000000000", 1),
                    word_cell("00 00", 1),
                    word_cell("0 0", 1),
                ],
                vec![
                    word_cell("0", 1),
                    word_cell_paragraphs(
                        &["000000000000000000000000000000", "00000000000000000000"],
                        2,
                    ),
                ],
            ],
        );
        assert_eq!(extent.column_widths, vec![60.0, 100.0, 100.0]);
        assert_eq!(paragraph.lines.len(), 2);
        let BlockExtent::Paragraph(spanning) = &extent.rows[1].cells[1].blocks[1] else {
            panic!()
        };
        assert_eq!(spanning.lines.len(), 1);
    }

    #[test]
    fn a_rotated_cell_keeps_its_column_while_another_widens_to_its_longest_word() {
        let word = 10.0 * 1139.0 / 128.0 + 2.0;
        let mut rotated = word_cell("0000 0000 0000 0000 0000 0000", 1);
        rotated["textDirection"] = json!("btLr");
        rotated["rowSpan"] = json!(2);
        let first = [word_cell("0000000000", 1), word_cell("00 00", 1), rotated];
        let second = [word_cell("0", 1), word_cell("0", 1)];
        let (extent, paragraph) = word_table_row_objects(
            None,
            vec![
                json!({"height":20,"cells":first}),
                json!({"height":20,"cells":second}),
            ],
        );
        assert_eq!(paragraph.lines.len(), 1);
        assert!((extent.column_widths[0] - word).abs() < 1e-3);
        assert!((extent.column_widths[1] - (160.0 - word)).abs() < 1e-3);
        assert_eq!(extent.column_widths[2], 100.0);
    }
}
