use std::collections::{BTreeMap, HashMap};

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::table_grid::{resolve_cell_grid, resolve_table_column_widths, resolve_table_width_px};
use crate::types::{
    BlockExtent, ChartExtent, ImageExtent, ImageRunPosition, LayoutBlock, ParagraphBlock,
    ParagraphExtent, Run, ShapeBlock, ShapeExtent, TableBlock, TableCellExtent, TableExtent,
    TableRowExtent, TextBoxBlock, TextBoxExtent,
};

const DEFAULT_CELL_PADDING_X: f64 = 7.0;
const DEFAULT_CELL_PADDING_Y: f64 = 0.0;
const ANCHOR_PROXIMITY: usize = 4;

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct FloatingZone {
    left_margin: f64,
    right_margin: f64,
    top_y: f64,
    bottom_y: f64,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    full_width_block: bool,
}

#[derive(Clone, Debug)]
struct AnchoredFloatingZone {
    zone: FloatingZone,
    anchor_block_index: usize,
    margin_relative: bool,
}

#[derive(Clone, Debug)]
pub struct FloatPageGeometry {
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
            false,
            false,
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

pub fn measure_blocks_with_floats(
    blocks: &mut [LayoutBlock],
    widths: &[f64],
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
) -> Result<Vec<BlockExtent>, String> {
    let default_width = widths.first().copied().unwrap_or(0.0);
    let extracted = extract_floating_zones(blocks, default_width, config, page_geometry)?;
    let mut margin_groups = BTreeMap::<u64, Vec<AnchoredFloatingZone>>::new();
    let mut paragraph_zones = Vec::new();
    for anchored in extracted {
        if anchored.margin_relative {
            margin_groups
                .entry(anchored.zone.top_y.to_bits())
                .or_default()
                .push(anchored);
        } else {
            paragraph_zones.push(anchored);
        }
    }
    let mut groups = group_overlapping_zones(paragraph_zones);
    groups.extend(margin_groups.into_values());
    let mut zones_by_anchor = HashMap::<usize, Vec<FloatingZone>>::new();
    for group in groups {
        let earliest = group
            .iter()
            .map(|anchored| anchored.anchor_block_index)
            .min()
            .unwrap_or(0);
        for anchored in group {
            let anchor = if anchored.zone.full_width_block && anchored.margin_relative {
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

    let mut cumulative_y = 0.0;
    let mut active_zones = Vec::new();
    let mut measured = Vec::with_capacity(blocks.len());
    for (index, block) in blocks.iter_mut().enumerate() {
        if let Some(zones) = zones_by_anchor.get(&index) {
            cumulative_y = 0.0;
            active_zones.clone_from(zones);
        }
        let width = widths.get(index).copied().unwrap_or(default_width);
        let extent = measure_block_with_context(
            block,
            width,
            config,
            (!active_zones.is_empty()).then_some(active_zones.as_slice()),
            cumulative_y,
        )?;
        if !matches!(block, LayoutBlock::Table(table) if table.floating.is_some()) {
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
    if !content_width.is_finite() || content_width <= 0.0 {
        return Ok(synthetic_paragraph_extent(paragraph, content_width));
    }
    let mut envelope = json!({
        "block": LayoutBlock::Paragraph(paragraph.clone()),
        "maxWidth": content_width,
        "fontChains": config.font_chains,
        "authoritativeShaping": config.authoritative_shaping,
    });
    let fields = envelope
        .as_object_mut()
        .expect("measurement envelope object");
    if !config.defaults.is_null() {
        fields.insert("defaults".to_owned(), config.defaults.clone());
    }
    if !config.compat.is_null() {
        fields.insert("compat".to_owned(), config.compat.clone());
    }
    if let Some(zones) = floating_zones {
        fields.insert(
            "floatingZones".to_owned(),
            serde_json::to_value(zones).expect("floating zones serialize"),
        );
        fields.insert("paragraphYOffset".to_owned(), json!(cumulative_y));
    }
    let Ok(extent) = crate::measure_paragraph_json_resident(&envelope.to_string()) else {
        return Ok(synthetic_paragraph_extent(paragraph, content_width));
    };
    serde_json::from_str(&extent)
        .map_err(|error| format!("parse paragraph extent: {error}"))
        .or_else(|_| Ok(synthetic_paragraph_extent(paragraph, content_width)))
}

fn synthetic_paragraph_extent(paragraph: &ParagraphBlock, content_width: f64) -> ParagraphExtent {
    let mut font_size = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.default_font_size)
        .unwrap_or(11.0);
    let mut character_count = 0;
    for run in &paragraph.runs {
        let Run::Text(text) = run else {
            continue;
        };
        character_count += text.text.encode_utf16().count();
        if let Some(size) = text.fmt.font_size.filter(|size| size.is_finite()) {
            font_size = font_size.max(size);
        }
    }
    if !font_size.is_finite() || font_size <= 0.0 {
        font_size = 11.0;
    }
    let font_size_px = font_size * 96.0 / 72.0;
    let line_height = font_size_px * 1.15;
    let tail_run = paragraph.runs.len().saturating_sub(1);
    let tail_char = paragraph.runs.get(tail_run).map_or(0, |run| match run {
        Run::Text(text) => text.text.encode_utf16().count(),
        _ => 0,
    });
    let spacing = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.spacing.as_ref());
    ParagraphExtent {
        lines: vec![crate::types::TypesetRow {
            head_run: 0,
            head_char: 0,
            tail_run,
            tail_char,
            width: if content_width.is_finite() && content_width > 0.0 {
                content_width.min(character_count as f64 * font_size_px * 0.5)
            } else {
                0.0
            },
            ascent: font_size_px * 0.8,
            descent: font_size_px * 0.2,
            line_height,
            ..crate::types::TypesetRow::default()
        }],
        total_height: spacing.and_then(|value| value.before).unwrap_or(0.0)
            + line_height
            + spacing.and_then(|value| value.after).unwrap_or(0.0),
    }
}

fn extract_floating_zones(
    blocks: &[LayoutBlock],
    content_width: f64,
    config: &MeasurementConfig,
    page_geometry: Option<&FloatPageGeometry>,
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
            _ => {}
        }
    }
    Ok(zones)
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
    let Some(floating) = &table.floating else {
        return Ok(());
    };
    let mut measured_table = table.clone();
    let measure = measure_table(&mut measured_table, content_width, config)?;
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
    let (left_margin, right_margin) = if x < content_width / 2.0 {
        clamp_margins(
            x + measure.total_width + floating.right_from_text.unwrap_or(12.0),
            0.0,
            content_width,
        )
    } else {
        clamp_margins(
            0.0,
            content_width - x + floating.left_from_text.unwrap_or(12.0),
            content_width,
        )
    };
    let top_y = floating.tblp_y.unwrap_or(0.0);
    zones.push(AnchoredFloatingZone {
        zone: FloatingZone {
            left_margin,
            right_margin,
            top_y: top_y - floating.top_from_text.unwrap_or(0.0),
            bottom_y: top_y + measure.total_height + floating.bottom_from_text.unwrap_or(0.0),
            full_width_block: false,
        },
        anchor_block_index: block_index,
        margin_relative: false,
    });
    Ok(())
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

fn group_overlapping_zones(zones: Vec<AnchoredFloatingZone>) -> Vec<Vec<AnchoredFloatingZone>> {
    let mut groups: Vec<Vec<AnchoredFloatingZone>> = Vec::new();
    for zone in zones {
        if let Some(group) = groups.iter_mut().find(|group| {
            group.iter().any(|other| {
                other.anchor_block_index.abs_diff(zone.anchor_block_index) <= ANCHOR_PROXIMITY
                    && zone.zone.top_y < other.zone.bottom_y
                    && zone.zone.bottom_y > other.zone.top_y
            })
        }) {
            group.push(zone);
        } else {
            groups.push(vec![zone]);
        }
    }
    groups
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 9_525.0
}

fn measure_shape(
    shape: &mut ShapeBlock,
    config: &MeasurementConfig,
) -> Result<ShapeExtent, String> {
    let inner_measures = shape
        .inner_text
        .as_ref()
        .map(|paragraphs| {
            paragraphs
                .iter()
                .map(|paragraph| measure_paragraph(paragraph, shape.width, config))
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

fn measure_table(
    table: &mut TableBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<TableExtent, String> {
    let explicit_width =
        resolve_table_width_px(table.width, table.width_type.as_deref(), content_width);
    let target_width = explicit_width.unwrap_or(content_width);
    let column_widths = resolve_table_column_widths(table, content_width);
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
            let measures = measure_blocks(
                &mut cell.blocks,
                (cell_width - left - right).max(1.0),
                config,
            )?;
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
        let mut max_border_height = 0.0_f64;
        for (cell_index, measured_cell) in measured_row.cells.iter_mut().enumerate() {
            let source_cell = &source_row.cells[cell_index];
            let mut content_height = 0.0_f64;
            let mut previous_after = 0.0_f64;
            for (block, measure) in source_cell.blocks.iter().zip(&measured_cell.blocks) {
                let visual = table_cell_block_height(block, measure);
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
            measured_cell.height = content_height
                + previous_after
                + source_cell
                    .padding
                    .as_ref()
                    .map_or(DEFAULT_CELL_PADDING_Y, |padding| padding.top)
                + source_cell
                    .padding
                    .as_ref()
                    .map_or(DEFAULT_CELL_PADDING_Y, |padding| padding.bottom);
            if source_cell.row_span.unwrap_or(1.0) <= 1.0 {
                max_height = max_height.max(measured_cell.height);
            }
            max_border_height = max_border_height.max(cell_border_height(source_cell));
        }
        exact[row_index] =
            source_row.height_rule.as_deref() == Some("exact") && source_row.height.is_some();
        measured_row.height = match (source_row.height, source_row.height_rule.as_deref()) {
            (Some(height), Some("exact")) => height,
            (Some(height), _) => (max_height + max_border_height).max(height),
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
            Run::Image(image) => Some(image.height),
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
    })
}

#[cfg(test)]
mod tests {
    use super::*;

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
    fn missing_font_chain_uses_the_reference_synthetic_extent() {
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

        assert_eq!(extent.lines[0].width, 32.0);
        assert_eq!(extent.lines[0].ascent, 12.8);
        assert_eq!(extent.lines[0].descent, 3.2);
        assert_eq!(extent.total_height, 23.4);
    }
}
