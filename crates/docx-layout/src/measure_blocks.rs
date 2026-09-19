use std::collections::{BTreeMap, HashMap};

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::cell_layout::{nested_table_float_offset, nested_table_horizontal_offset};
use crate::floating_objects::MIN_WRAP_SEGMENT_WIDTH;
use crate::table_grid::{
    content_sized_columns, count_table_columns, grow_content_sized_columns, resolve_cell_grid,
    resolve_table_column_widths, resolve_table_width_px,
};
use crate::types::{
    BlockExtent, ChartExtent, FloatingTablePosition, ImageExtent, ImageRunPosition, LayoutBlock,
    ParagraphBlock, ParagraphExtent, ParagraphSpacing, Run, ShapeBlock, ShapeExtent, TableBlock,
    TableCellExtent, TableExtent, TableRowExtent, TextBoxBlock, TextBoxExtent, TypesetRow,
};
use ooxml_text::{LineBox, LineSpacingRule, apply_spacing_rule};

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

pub fn collect_font_requirements(blocks: &[LayoutBlock]) -> Vec<FontRequirement> {
    let mut requirements = BTreeMap::<String, FontRequirement>::new();
    walk_paragraphs(blocks, &mut |paragraph| {
        let scripts = paragraph_scripts(paragraph);
        collect_paragraph_font_requirements(paragraph, &scripts, &mut requirements);
    });
    requirements.into_values().collect()
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
) {
    let key = format!(
        "{}|{}|{}",
        family.to_lowercase(),
        u8::from(bold),
        u8::from(italic)
    );
    let requirement = requirements
        .entry(key.clone())
        .or_insert_with(|| FontRequirement {
            key,
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
}

fn collect_paragraph_font_requirements(
    paragraph: &ParagraphBlock,
    scripts: &[String],
    requirements: &mut BTreeMap<String, FontRequirement>,
) {
    let default_family = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.default_font_family.as_deref())
        .unwrap_or("Calibri");
    add_font_requirement(default_family, false, false, scripts, requirements);
    for run in &paragraph.runs {
        let (formatting, include_regular) = match run {
            Run::Text(text) => (&text.fmt, true),
            Run::Tab(tab) => (&tab.fmt, false),
            Run::Field(field) => (&field.fmt, false),
            _ => continue,
        };
        let family = formatting.font_family.as_deref().unwrap_or(default_family);
        let bold = formatting.bold.unwrap_or(false);
        let italic = formatting.italic.unwrap_or(false);
        add_font_requirement(family, bold, italic, scripts, requirements);
        if include_regular {
            add_font_requirement(family, false, false, scripts, requirements);
        }
        let Some(slots) = &formatting.font_slots else {
            continue;
        };
        for family in [
            slots.ascii.as_deref(),
            slots.h_ansi.as_deref(),
            slots.east_asia.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            add_font_requirement(family, bold, italic, scripts, requirements);
            if include_regular {
                add_font_requirement(family, false, false, scripts, requirements);
            }
        }
        if let Some(family) = slots.cs.as_deref() {
            add_font_requirement(
                family,
                formatting.bold_cs.unwrap_or(bold),
                formatting.italic_cs.unwrap_or(italic),
                scripts,
                requirements,
            );
            if include_regular {
                add_font_requirement(family, false, false, scripts, requirements);
            }
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
        add_font_requirement(
            attrs
                .list_marker_font_family
                .as_deref()
                .or(first_run_family)
                .unwrap_or(default_family),
            attrs.list_marker_bold.unwrap_or(false),
            attrs.list_marker_italic.unwrap_or(false),
            scripts,
            requirements,
        );
    }
}

fn paragraph_scripts(paragraph: &ParagraphBlock) -> Vec<String> {
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
    if han && !kana && !hangul {
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
    let default_width = widths.first().copied().unwrap_or(0.0);
    let extracted =
        extract_floating_zones(blocks, default_width, config, page_geometry, shape_offsets)?;
    let mut margin_groups = BTreeMap::<u64, Vec<AnchoredFloatingZone>>::new();
    let mut paragraph_zones = BTreeMap::<usize, Vec<FloatingZone>>::new();
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
    let mut zones_by_anchor = HashMap::<usize, Vec<FloatingZone>>::new();
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

    let section_break_marks = blocks
        .iter()
        .enumerate()
        .map(|(index, block)| {
            matches!(block, LayoutBlock::Paragraph(paragraph) if paragraph.runs.is_empty())
                && matches!(blocks.get(index + 1), Some(LayoutBlock::SectionBreak(_)))
        })
        .collect::<Vec<_>>();

    let mut cumulative_y = 0.0;
    let mut active_zones = Vec::new();
    let mut measured = Vec::with_capacity(blocks.len());
    for (index, block) in blocks.iter_mut().enumerate() {
        if matches!(
            block,
            LayoutBlock::PageBreak(_) | LayoutBlock::ColumnBreak(_) | LayoutBlock::SectionBreak(_)
        ) || crate::keep_together::paragraph_breaks_before(block)
        {
            active_zones.clear();
            cumulative_y = 0.0;
        }
        if let Some(zones) = paragraph_zones.get(&index) {
            // A paragraph-anchored band hangs off its own anchor, which sits at
            // `cumulative_y` in the frame the earlier bands were measured in.
            if active_zones.len() + zones.len() <= MAX_ACTIVE_ZONES {
                active_zones.extend(zones.iter().map(|zone| FloatingZone {
                    top_y: zone.top_y + cumulative_y,
                    bottom_y: zone.bottom_y + cumulative_y,
                    ..zone.clone()
                }));
            } else {
                cumulative_y = 0.0;
                active_zones.clone_from(zones);
            }
        }
        if let Some(zones) = zones_by_anchor.get(&index) {
            // Anchors the flow has not advanced past share one origin.
            if cumulative_y == 0.0 && active_zones.len() + zones.len() <= MAX_ACTIVE_ZONES {
                active_zones.extend(zones.iter().cloned());
            } else {
                cumulative_y = 0.0;
                active_zones.clone_from(zones);
            }
        }
        let width = widths.get(index).copied().unwrap_or(default_width);
        // A bare paragraph mark carrying section properties is the section
        // break itself; Word prints no line for it.
        let extent = if section_break_marks[index] {
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
                cumulative_y,
            )?
        };
        if !matches!(block, LayoutBlock::Table(table) if table.floating.is_some())
            && !matches!(block, LayoutBlock::Shape(shape) if anchored_shape(shape))
        {
            cumulative_y += extent_height(&extent);
        }
        measured.push(extent);
    }
    Ok(measured)
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
    let mut extent = if !content_width.is_finite() || content_width <= 0.0 {
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
    };
    measure_horizontal_rules(paragraph, &mut extent);
    Ok(extent)
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
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
    shape_offsets: &BTreeMap<usize, f64>,
) -> Result<Vec<AnchoredFloatingZone>, String> {
    let mut zones = Vec::new();
    for (block_index, block) in blocks.iter().enumerate() {
        match block {
            LayoutBlock::Paragraph(paragraph) => {
                extract_image_zones(paragraph, block_index, content_width, &mut zones);
            }
            LayoutBlock::Table(table) => {
                extract_table_zone(table, block_index, content_width, config, &mut zones)?;
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
    }
    Ok(zones)
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

fn extract_table_zone(
    table: &TableBlock,
    block_index: usize,
    content_width: f64,
    config: &MeasurementConfig,
    zones: &mut Vec<AnchoredFloatingZone>,
) -> Result<(), String> {
    if table.floating.is_none() {
        return Ok(());
    }
    let mut measured_table = table.clone();
    let measure = measure_table(&mut measured_table, content_width, config)?;
    if let Some(mut zone) = table_floating_zone(table, &measure, content_width) {
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
    ))
}

fn table_floating_zone_at_x(
    floating: &FloatingTablePosition,
    measure: &TableExtent,
    content_width: f64,
    x: f64,
) -> FloatingZone {
    let (left_margin, right_margin) = if x < content_width / 2.0 {
        (
            x + measure.total_width + floating.right_from_text.unwrap_or(12.0),
            0.0,
        )
    } else {
        (
            0.0,
            content_width - x + floating.left_from_text.unwrap_or(12.0),
        )
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
            let mut zone = table_floating_zone_at_x(floating, measure, content_width, x);
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

fn measure_table(
    table: &mut TableBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<TableExtent, String> {
    let explicit_width =
        resolve_table_width_px(table.width, table.width_type.as_deref(), content_width);
    let target_width = explicit_width.unwrap_or(content_width);
    let mut column_widths = resolve_table_column_widths(table, content_width);
    let content_sized = content_sized_columns(table, content_width, &column_widths);
    if !content_sized.is_empty() {
        let maximums = column_content_maximums(table, &content_sized, content_width, config)?;
        grow_content_sized_columns(table, content_width, &maximums, &mut column_widths);
    }
    let grid = resolve_cell_grid(table);
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
            let rotated = matches!(cell.text_direction.as_deref(), Some("btLr" | "tbRl"));
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
            if matches!(source_cell.text_direction.as_deref(), Some("btLr" | "tbRl")) {
                content_height = measured_cell
                    .blocks
                    .iter()
                    .filter_map(|measure| {
                        if let BlockExtent::Paragraph(paragraph) = measure {
                            Some(
                                paragraph
                                    .lines
                                    .iter()
                                    .map(|line| line.width)
                                    .fold(0.0, f64::max),
                            )
                        } else {
                            None
                        }
                    })
                    .fold(0.0, f64::max);
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

    let natural: Vec<f64> = rows.iter().map(|row| row.height).collect();
    for row_index in 0..rows.len() {
        for cell_index in 0..table.rows[row_index].cells.len() {
            let source_cell = &table.rows[row_index].cells[cell_index];
            let row_span = source_cell.row_span.unwrap_or(1.0).max(1.0) as usize;
            if row_span <= 1 {
                continue;
            }
            let last = (row_index + row_span - 1).min(rows.len() - 1);
            let needed = rows[row_index].cells[cell_index].height + cell_border_height(source_cell);
            let spanned = natural[row_index..=last].iter().sum::<f64>();
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
    spacing.and_then(|value| value.before).unwrap_or(0.0)
        + image_height
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

    #[test]
    fn a_bare_section_break_paragraph_mark_takes_no_line() {
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
            measure_blocks_with_floats(&mut blocks, &[200.0; 3], &config, None).unwrap()
        };
        let empty = json!({"kind":"paragraph","id":"mark","runs":[]});
        let spaced = json!({"kind":"paragraph","id":"mark","runs":[{"kind":"text","text":" "}]});
        let section = json!({"kind":"sectionBreak","id":"sect:mark"});
        let tail = json!({"kind":"paragraph","id":"tail","runs":[{"kind":"text","text":"words"}]});

        let marked = measure(json!([empty, section, tail]));
        let BlockExtent::Paragraph(mark) = &marked[0] else {
            panic!()
        };
        assert!(mark.lines.is_empty());
        assert_eq!(mark.total_height, 0.0);

        for kept in [
            measure(json!([empty, tail])),
            measure(json!([spaced, section, tail])),
        ] {
            let BlockExtent::Paragraph(mark) = &kept[0] else {
                panic!()
            };
            assert_eq!(mark.lines.len(), 1);
            assert!(mark.total_height > 0.0);
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
    fn collapsed_borders_expand_minimum_and_auto_rows_but_not_exact_rows() {
        for (height, rule, expected) in [
            (Some(40.0), "atLeast", 41.0),
            (Some(40.0), "exact", 40.0),
            (None, "auto", 17.0),
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
    fn minimum_row_height_reserves_cell_margins_outside_the_content_minimum() {
        for (minimum, content, expected) in [(40, 16, 56.0), (80, 16, 96.0), (40, 60, 76.0)] {
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

        let requirements = collect_font_requirements(&blocks);
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
}
