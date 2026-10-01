use std::borrow::Cow;

use serde::Serialize;

use crate::measure_blocks::{MeasurementConfig, extent_height, measure_blocks, measure_paragraph};
use crate::paragraph_spacing::apply_contextual_spacing_blocks;
use crate::types::{
    AxisPosition, BlockExtent, BlockId, BoxEdges, FieldRun, ImageRun, ImageRunPosition, Layout,
    LayoutBlock, MeasuredBlock, PageFloatBand, PageMargins, ParagraphBlock, Run, Size,
};

const DEFAULT_HF_DISTANCE_PX: f64 = 48.0;
const MIN_CONTENT_HEIGHT_PX: f64 = 24.0;

#[derive(Default)]
pub(crate) struct HeaderFooterFlow {
    pub cursor: f64,
    after: f64,
}

impl HeaderFooterFlow {
    pub fn place(&mut self, height: f64, before: f64, after: f64) -> f64 {
        let y = self.cursor + self.after.max(before);
        self.cursor = y + height;
        self.after = after;
        y
    }

    pub fn height(&self) -> f64 {
        self.cursor + self.after
    }
}

fn block_spacing(block: &LayoutBlock) -> (f64, f64) {
    let spacing = match block {
        LayoutBlock::Paragraph(paragraph) => paragraph
            .attrs
            .as_ref()
            .and_then(|attrs| attrs.spacing.as_ref()),
        _ => None,
    };
    (
        spacing.and_then(|s| s.before).unwrap_or(0.0),
        spacing.and_then(|s| s.after).unwrap_or(0.0),
    )
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum HeaderFooterKind {
    Header,
    Footer,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum HeaderFooterType {
    Default,
    First,
    Even,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterVariant {
    pub r_id: String,
    pub kind: HeaderFooterKind,
    #[serde(rename = "type")]
    pub hf_type: HeaderFooterType,
    pub section_index: usize,
    pub measured: Vec<MeasuredBlock>,
    pub height: f64,
    pub flow_height: f64,
    pub visual_top: f64,
    pub visual_bottom: f64,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub field_widths: Vec<HeaderFooterFieldWidths>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterFieldWidths {
    pub pm_start: i64,
    pub fallback_width: f64,
    pub per_page: Vec<f64>,
}

#[derive(Clone, Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterPayload {
    pub title_pg: bool,
    pub even_and_odd_headers: bool,
    pub title_page_sections: Vec<usize>,
    pub even_and_odd_sections: Vec<usize>,
    pub variants: Vec<HeaderFooterVariant>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub watermark: Option<serde_json::Value>,
}

#[derive(Clone, Copy)]
pub struct HeaderFooterMetrics<'a> {
    pub kind: HeaderFooterKind,
    pub page_size: &'a Size,
    pub margins: &'a PageMargins,
}

pub fn measure_header_footer(
    r_id: String,
    kind: HeaderFooterKind,
    hf_type: HeaderFooterType,
    section_index: usize,
    blocks: Vec<LayoutBlock>,
    content_width: f64,
    metrics: HeaderFooterMetrics<'_>,
    config: &MeasurementConfig,
) -> Result<Option<HeaderFooterVariant>, String> {
    if blocks.is_empty() {
        return Ok(None);
    }
    let mut blocks = blocks;
    apply_contextual_spacing_blocks(&mut blocks);
    let mut detached = float_detached_top_and_bottom_images(&mut blocks);
    let mut measures = measure_blocks(&mut blocks, content_width, config)?;
    if restore_overlapping_detached_images(&mut blocks, &measures, &mut detached, metrics) {
        measures = measure_blocks(&mut blocks, content_width, config)?;
    }
    let height = measures.iter().map(extent_height).sum();
    let mut flow = HeaderFooterFlow::default();
    for (block, measure) in blocks.iter().zip(&measures) {
        if contributes_to_flow(block) {
            let (before, after) = block_spacing(block);
            flow.place(
                (extent_height(measure) - before - after).max(0.0),
                before,
                after,
            );
        }
    }
    let flow_height = flow.height();
    let (visual_top, visual_bottom) = visual_bounds(&blocks, &measures, flow_height, metrics);
    let measured = blocks
        .into_iter()
        .zip(measures)
        .map(|(block, measure)| MeasuredBlock { block, measure })
        .collect();
    Ok(Some(HeaderFooterVariant {
        r_id,
        kind,
        hf_type,
        section_index,
        measured,
        height,
        flow_height,
        visual_top,
        visual_bottom,
        field_widths: Vec::new(),
    }))
}

/// Floats detached images, retaining their previous modes for overlap checks.
fn float_detached_top_and_bottom_images(
    blocks: &mut [LayoutBlock],
) -> Vec<(usize, usize, Option<String>)> {
    let mut detached_images = Vec::new();
    for (block_index, block) in blocks.iter_mut().enumerate() {
        let LayoutBlock::Paragraph(paragraph) = block else {
            continue;
        };
        for (run_index, run) in paragraph.runs.iter_mut().enumerate() {
            let Run::Image(image) = run else {
                continue;
            };
            let detached = image
                .position
                .as_ref()
                .and_then(|position| position.vertical.as_ref())
                .and_then(|vertical| vertical.relative_to.as_deref())
                .is_some_and(|relative| !matches!(relative, "paragraph" | "line"));
            if detached
                && image
                    .position
                    .as_ref()
                    .and_then(|position| position.behind_doc)
                    != Some(true)
                && image.wrap_type.as_deref() == Some("topAndBottom")
                && image.display_mode.as_deref() != Some("float")
            {
                let previous = image.display_mode.replace("float".to_owned());
                detached_images.push((block_index, run_index, previous));
            }
        }
    }
    if detached_images.len() > 16 {
        restore_detached_images(blocks, &mut detached_images);
    }
    detached_images
}

fn restore_detached_images(
    blocks: &mut [LayoutBlock],
    detached: &mut Vec<(usize, usize, Option<String>)>,
) {
    for (block_index, run_index, previous) in detached.drain(..) {
        let LayoutBlock::Paragraph(paragraph) = &mut blocks[block_index] else {
            unreachable!();
        };
        let Run::Image(image) = &mut paragraph.runs[run_index] else {
            unreachable!();
        };
        image.display_mode = previous;
    }
}

fn restore_overlapping_detached_images(
    blocks: &mut [LayoutBlock],
    measures: &[BlockExtent],
    detached: &mut Vec<(usize, usize, Option<String>)>,
    metrics: HeaderFooterMetrics<'_>,
) -> bool {
    if detached.is_empty() {
        return false;
    }
    let unmeasured_text = blocks.iter().any(|block| {
        matches!(block, LayoutBlock::TextBox(_))
            || (!contributes_to_flow(block)
                && matches!(block, LayoutBlock::Table(_) | LayoutBlock::Shape(_)))
    });
    let mut flow = HeaderFooterFlow::default();
    let bounds: Vec<_> = blocks
        .iter()
        .zip(measures)
        .map(|(block, measure)| {
            if !contributes_to_flow(block) {
                return None;
            }
            let (before, after) = block_spacing(block);
            let height = (extent_height(measure) - before - after).max(0.0);
            let top = flow.place(height, before, after);
            Some((top, top + height))
        })
        .collect();
    let flow_height = flow.height();
    let distance = match metrics.kind {
        HeaderFooterKind::Header => metrics.margins.header,
        HeaderFooterKind::Footer => metrics.margins.footer,
    }
    .unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let flow_top = match metrics.kind {
        HeaderFooterKind::Header => distance,
        HeaderFooterKind::Footer => metrics.page_size.h - distance - flow_height,
    };
    let geom = float_geometry(metrics);
    let overlaps = detached.iter().any(|(block_index, run_index, _)| {
        let LayoutBlock::Paragraph(paragraph) = &blocks[*block_index] else {
            return false;
        };
        let Run::Image(image) = &paragraph.runs[*run_index] else {
            return false;
        };
        let (x, y) = crate::display_list::resolve_anchored_position(
            image.position.as_ref(),
            image.css_float.as_deref(),
            image.width,
            image.height,
            flow_top - geom.margin_top,
            &geom,
        );
        let top = crate::display_list::clamp_wrapped_float_y(
            geom.margin_top + y,
            image.height,
            image.wrap_type.as_deref(),
            geom.page_height,
        );
        let bottom = top + image.height + image.dist_bottom.unwrap_or(0.0).max(0.0);
        let top = top - image.dist_top.unwrap_or(0.0).max(0.0);
        unmeasured_text
            || image.rotation_bounds.is_some()
            || image.inline_shape.is_some()
            || !image.width.is_finite()
            || image.width <= 0.0
            || image.height <= 0.0
            || !x.is_finite()
            || !top.is_finite()
            || !bottom.is_finite()
            || bounds
                .iter()
                .flatten()
                .any(|&(y, end)| flow_top + y < bottom && flow_top + end > top)
    });
    if overlaps {
        restore_detached_images(blocks, detached);
    }
    overlaps
}

pub fn resolve_header_footer_field_widths(
    payload: &mut HeaderFooterPayload,
    layout: &Layout,
    config: &MeasurementConfig,
) -> Result<(), String> {
    let total_pages = if layout.partial {
        String::new()
    } else {
        layout.pages.len().to_string()
    };
    for variant in &mut payload.variants {
        let mut widths = Vec::new();
        for measured in &variant.measured {
            let mut fields = Vec::new();
            page_fields(&measured.block, &mut fields);
            for field in fields {
                let Some(pm_start) = integral_position(field.pm_start) else {
                    continue;
                };
                let fallback = field
                    .fallback
                    .as_deref()
                    .filter(|value| !value.is_empty())
                    .unwrap_or("1");
                let fallback_width = measure_field_text(field, fallback, config)?;
                let per_page = layout
                    .pages
                    .iter()
                    .map(|page| {
                        let text: Cow<'_, str> = if field.field_type == "NUMPAGES" {
                            Cow::Borrowed(total_pages.as_str())
                        } else {
                            crate::regions::page_field_text(
                                page.page_label.as_deref(),
                                u64::from(page.number),
                            )
                        };
                        if layout.partial && field.field_type == "NUMPAGES" {
                            match field.fallback.as_deref().filter(|value| !value.is_empty()) {
                                Some(text) if layout.cached_page_totals => {
                                    measure_field_text(field, text, config)
                                }
                                _ => Ok(0.0),
                            }
                        } else {
                            measure_field_text(field, &text, config)
                        }
                    })
                    .collect::<Result<Vec<_>, _>>()?;
                widths.push(HeaderFooterFieldWidths {
                    pm_start,
                    fallback_width,
                    per_page,
                });
            }
        }
        variant.field_widths = widths;
    }
    Ok(())
}

/// The PAGE and NUMPAGES fields of a block, table cells included.
fn page_fields<'a>(block: &'a LayoutBlock, fields: &mut Vec<&'a FieldRun>) {
    match block {
        LayoutBlock::Paragraph(paragraph) => {
            fields.extend(paragraph.runs.iter().filter_map(|run| match run {
                Run::Field(field) if matches!(field.field_type.as_str(), "PAGE" | "NUMPAGES") => {
                    Some(field)
                }
                _ => None,
            }));
        }
        LayoutBlock::Table(table) => {
            for cell in table.rows.iter().flat_map(|row| &row.cells) {
                for block in &cell.blocks {
                    page_fields(block, fields);
                }
            }
        }
        _ => {}
    }
}

fn integral_position(value: Option<f64>) -> Option<i64> {
    let value = value?;
    (value.is_finite()
        && value.fract() == 0.0
        && value >= i64::MIN as f64
        && value <= i64::MAX as f64)
        .then_some(value as i64)
}

fn measure_field_text(
    field: &FieldRun,
    text: &str,
    config: &MeasurementConfig,
) -> Result<f64, String> {
    let paragraph = ParagraphBlock {
        sdt_groups: None,
        id: BlockId::Num(0.0),
        para_id: None,
        runs: vec![Run::Field(FieldRun {
            fmt: field.fmt.clone(),
            field_type: field.field_type.clone(),
            raw_type: None,
            instruction: None,
            fallback: Some(text.to_owned()),
            locked: false,
            nested_sequences: Vec::new(),
            pm_start: None,
            pm_end: None,
        })],
        attrs: None,
        pm_start: None,
        pm_end: None,
    };
    let extent = measure_paragraph(&paragraph, 1_000_000.0, config)?;
    Ok(extent.lines.first().map_or(0.0, |line| line.width))
}

pub fn contributes_to_flow(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(_) => true,
        LayoutBlock::Table(table) => table.floating.is_none(),
        LayoutBlock::Image(image) => {
            image.anchor.as_ref().and_then(|anchor| anchor.is_anchored) != Some(true)
        }
        LayoutBlock::Shape(shape) => shape.position.is_none(),
        LayoutBlock::Chart(_) => true,
        LayoutBlock::TextBox(text_box) => {
            matches!(text_box.display_mode.as_deref(), None | Some("inline"))
        }
        _ => false,
    }
}

fn visual_bounds(
    blocks: &[LayoutBlock],
    measures: &[BlockExtent],
    height: f64,
    metrics: HeaderFooterMetrics<'_>,
) -> (f64, f64) {
    let mut visual_top = 0.0_f64;
    let mut visual_bottom = 0.0_f64;
    let mut flow = HeaderFooterFlow::default();
    for (block, measure) in blocks.iter().zip(measures) {
        let (before, after) = block_spacing(block);
        let block_height = (extent_height(measure) - before - after).max(0.0);
        let anchor_y = flow.cursor;
        let cursor = if contributes_to_flow(block) {
            flow.place(block_height, before, after)
        } else {
            flow.cursor
        };
        match block {
            LayoutBlock::Paragraph(paragraph) => {
                visual_top = visual_top.min(cursor);
                visual_bottom = visual_bottom.max(cursor + block_height);
                for run in &paragraph.runs {
                    let Run::Image(image) = run else {
                        continue;
                    };
                    if image.position.is_none() {
                        continue;
                    }
                    let top = image_visual_top(image, anchor_y, height, metrics);
                    visual_top = visual_top.min(top);
                    visual_bottom = visual_bottom.max(top + image.height);
                }
            }
            LayoutBlock::TextBox(_) => {
                visual_top = visual_top.min(cursor);
                visual_bottom = visual_bottom.max(cursor + block_height);
            }
            LayoutBlock::Shape(shape) if shape.position.is_some() => {
                let distance = match metrics.kind {
                    HeaderFooterKind::Header => metrics.margins.header,
                    HeaderFooterKind::Footer => metrics.margins.footer,
                }
                .unwrap_or(DEFAULT_HF_DISTANCE_PX);
                let flow_top = match metrics.kind {
                    HeaderFooterKind::Header => distance,
                    HeaderFooterKind::Footer => metrics.page_size.h - distance - height,
                };
                let (_, top) = crate::anchor::resolve_position(
                    shape.position.as_ref(),
                    shape.width,
                    shape.height,
                    &crate::anchor::AnchorFrame {
                        page_width: metrics.page_size.w,
                        page_height: metrics.page_size.h,
                        margin_left: metrics.margins.left,
                        margin_right: metrics.margins.right,
                        margin_top: metrics.margins.top,
                        margin_bottom: metrics.margins.bottom,
                        flow_x: metrics.margins.left,
                        flow_y: flow_top + cursor,
                        flow_width: metrics.page_size.w
                            - metrics.margins.left
                            - metrics.margins.right,
                        flow_height: 0.0,
                        odd_page: true,
                    },
                );
                visual_top = visual_top.min(top - flow_top);
                visual_bottom = visual_bottom.max(top - flow_top + block_height);
            }
            LayoutBlock::Table(_)
            | LayoutBlock::Image(_)
            | LayoutBlock::Shape(_)
            | LayoutBlock::Chart(_) => {
                visual_top = visual_top.min(cursor);
                visual_bottom = visual_bottom.max(cursor + block_height);
            }
            _ => {}
        }
    }
    (visual_top, visual_bottom.max(flow.height()))
}

fn image_visual_top(
    image: &ImageRun,
    paragraph_y: f64,
    flow_height: f64,
    metrics: HeaderFooterMetrics<'_>,
) -> f64 {
    let distance = match metrics.kind {
        HeaderFooterKind::Header => metrics.margins.header.unwrap_or(DEFAULT_HF_DISTANCE_PX),
        HeaderFooterKind::Footer => metrics.margins.footer.unwrap_or(DEFAULT_HF_DISTANCE_PX),
    };
    let flow_top = match metrics.kind {
        HeaderFooterKind::Header => distance,
        HeaderFooterKind::Footer => metrics.page_size.h - distance - flow_height,
    };
    let Some(vertical) = image
        .position
        .as_ref()
        .and_then(|position| position.vertical.as_ref())
    else {
        return paragraph_y;
    };
    let offset = vertical.pos_offset.map(emu_to_pixels);
    match vertical.relative_to.as_deref() {
        Some("page") => {
            if let Some(offset) = offset {
                return offset - flow_top;
            }
            match vertical.align.as_deref() {
                Some("top") => -flow_top,
                Some("bottom") => metrics.page_size.h - image.height - flow_top,
                Some("center") => (metrics.page_size.h - image.height) / 2.0 - flow_top,
                _ => paragraph_y,
            }
        }
        Some("margin") => {
            let margin_height = metrics.page_size.h - metrics.margins.top - metrics.margins.bottom;
            if let Some(offset) = offset {
                return metrics.margins.top + offset - flow_top;
            }
            match vertical.align.as_deref() {
                Some("top") => metrics.margins.top - flow_top,
                Some("bottom") => metrics.margins.top + margin_height - image.height - flow_top,
                Some("center") => {
                    metrics.margins.top + (margin_height - image.height) / 2.0 - flow_top
                }
                _ => paragraph_y,
            }
        }
        _ => offset.map_or(paragraph_y, |offset| paragraph_y + offset),
    }
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 914_400.0 * 96.0
}

fn float_geometry(metrics: HeaderFooterMetrics<'_>) -> crate::display_list::PageFloatGeom {
    crate::display_list::PageFloatGeom {
        page_width: metrics.page_size.w,
        page_height: metrics.page_size.h,
        margin_left: metrics.margins.left,
        margin_top: metrics.margins.top.abs(),
        content_width: metrics.page_size.w - metrics.margins.left - metrics.margins.right,
        content_height: metrics.page_size.h
            - metrics.margins.top.abs()
            - metrics.margins.bottom.abs(),
    }
}

pub fn header_footer_float_bands(
    variant: &HeaderFooterVariant,
    metrics: HeaderFooterMetrics<'_>,
) -> Vec<PageFloatBand> {
    let distance = match metrics.kind {
        HeaderFooterKind::Header => metrics.margins.header,
        HeaderFooterKind::Footer => metrics.margins.footer,
    }
    .unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let flow_top = match metrics.kind {
        HeaderFooterKind::Header => distance,
        HeaderFooterKind::Footer => metrics.page_size.h - distance - variant.flow_height,
    };
    let geom = float_geometry(metrics);
    let mut bands = Vec::new();
    let mut add = |position: Option<&ImageRunPosition>,
                   size: Size,
                   wrap: Option<&str>,
                   behind: bool,
                   distances: BoxEdges,
                   anchor_y: f64,
                   emu: bool,
                   image: bool,
                   css_float: Option<&str>| {
        let Some(position) = position else {
            return;
        };
        if behind
            || position.behind_doc == Some(true)
            || !matches!(wrap, Some("topAndBottom" | "square" | "tight" | "through"))
            || !size.w.is_finite()
            || !size.h.is_finite()
            || size.w <= 0.0
            || size.h <= 0.0
        {
            return;
        }
        let left = metrics.margins.left;
        let right = metrics.page_size.w - metrics.margins.right;
        let mut frame = crate::anchor::AnchorFrame {
            page_width: metrics.page_size.w,
            page_height: metrics.page_size.h,
            margin_left: left,
            margin_right: metrics.margins.right,
            margin_top: metrics.margins.top.abs(),
            margin_bottom: metrics.margins.bottom.abs(),
            flow_x: left,
            flow_y: flow_top + anchor_y,
            flow_width: right - left,
            flow_height: 0.0,
            odd_page: true,
        };
        let finite = |value: f64| {
            if value.is_finite() {
                value.max(0.0)
            } else {
                0.0
            }
        };
        for odd_page in [true, false] {
            frame.odd_page = odd_page;
            let (x, y) = if emu && image {
                let (x, y) = crate::display_list::resolve_anchored_position(
                    Some(position),
                    css_float,
                    size.w,
                    size.h,
                    flow_top + anchor_y - geom.margin_top,
                    &geom,
                );
                (geom.margin_left + x, geom.margin_top + y)
            } else if emu {
                crate::display_list::resolve_hf_box_position(
                    Some(position),
                    css_float,
                    size.w,
                    size.h,
                    flow_top + anchor_y - geom.margin_top,
                    &geom,
                )
            } else {
                crate::anchor::resolve_position(Some(position), size.w, size.h, &frame)
            };
            let y = if image {
                crate::display_list::clamp_wrapped_float_y(y, size.h, wrap, geom.page_height)
            } else {
                y
            };
            let full_width = x - finite(distances.left) - left
                < crate::floating_objects::MIN_WRAP_SEGMENT_WIDTH
                && right - x - size.w - finite(distances.right)
                    < crate::floating_objects::MIN_WRAP_SEGMENT_WIDTH;
            if y.is_finite() && (wrap == Some("topAndBottom") || full_width) {
                bands.push(PageFloatBand {
                    top: y - finite(distances.top),
                    bottom: y + size.h + finite(distances.bottom),
                    odd_page: Some(odd_page),
                });
            }
        }
    };
    let mut flow = HeaderFooterFlow::default();
    for measured in &variant.measured {
        let block = &measured.block;
        let anchor_y = flow.cursor;
        let (before, after) = block_spacing(block);
        let height = (extent_height(&measured.measure) - before - after).max(0.0);
        if contributes_to_flow(block) {
            flow.place(height, before, after);
        }
        let zero = || BoxEdges {
            top: 0.0,
            right: 0.0,
            bottom: 0.0,
            left: 0.0,
        };
        match block {
            LayoutBlock::Paragraph(paragraph) => {
                for run in &paragraph.runs {
                    if let Run::Image(image) = run {
                        let bound = |axis: &str| {
                            image
                                .rotation_bounds
                                .as_ref()
                                .and_then(|bounds| bounds.get(axis))
                                .and_then(serde_json::Value::as_f64)
                        };
                        add(
                            image.position.as_ref(),
                            Size {
                                w: bound("width").unwrap_or(image.width),
                                h: bound("height").unwrap_or(image.height),
                            },
                            image.wrap_type.as_deref(),
                            false,
                            BoxEdges {
                                top: image.dist_top.unwrap_or(0.0),
                                right: image.dist_right.unwrap_or(0.0),
                                bottom: image.dist_bottom.unwrap_or(0.0),
                                left: image.dist_left.unwrap_or(0.0),
                            },
                            anchor_y,
                            true,
                            true,
                            image.css_float.as_deref(),
                        );
                    }
                }
            }
            LayoutBlock::Shape(shape) => add(
                shape.position.as_ref(),
                Size {
                    w: shape.width,
                    h: shape.height,
                },
                shape.wrap_type.as_deref(),
                shape.behind_doc == Some(true),
                shape.wrap_distances.clone().unwrap_or_else(zero),
                anchor_y,
                false,
                false,
                None,
            ),
            LayoutBlock::TextBox(text_box) => add(
                text_box.position.as_ref(),
                Size {
                    w: text_box.width,
                    h: height,
                },
                text_box.wrap_type.as_deref(),
                false,
                BoxEdges {
                    top: text_box.dist_top.unwrap_or(0.0),
                    right: text_box.dist_right.unwrap_or(0.0),
                    bottom: text_box.dist_bottom.unwrap_or(0.0),
                    left: text_box.dist_left.unwrap_or(0.0),
                },
                anchor_y,
                true,
                false,
                text_box.css_float.as_deref(),
            ),
            LayoutBlock::Image(image) => {
                if let Some(anchor) = &image.anchor {
                    add(
                        anchor.position.as_ref(),
                        Size {
                            w: image.width,
                            h: image.height,
                        },
                        anchor.wrap_type.as_deref(),
                        anchor.behind_doc == Some(true),
                        zero(),
                        anchor_y,
                        false,
                        true,
                        None,
                    );
                }
            }
            LayoutBlock::Table(table) => {
                if let Some(floating) = &table.floating {
                    let axis = |relative: Option<&str>, offset, align: Option<&str>, fallback| {
                        Some(AxisPosition {
                            relative_to: Some(relative.unwrap_or(fallback).to_owned()),
                            pos_offset: offset,
                            align: align.map(str::to_owned),
                        })
                    };
                    let position = ImageRunPosition {
                        horizontal: axis(
                            floating.horz_anchor.as_deref(),
                            floating.tblp_x,
                            floating.tblp_x_spec.as_deref(),
                            "margin",
                        ),
                        vertical: axis(
                            floating.vert_anchor.as_deref(),
                            floating.tblp_y,
                            floating.tblp_y_spec.as_deref(),
                            "paragraph",
                        ),
                        use_simple_pos: None,
                        simple_pos: None,
                        relative_height: None,
                        behind_doc: None,
                    };
                    if let BlockExtent::Table(extent) = &measured.measure {
                        add(
                            Some(&position),
                            Size {
                                w: extent.total_width,
                                h: extent.total_height,
                            },
                            Some("square"),
                            false,
                            BoxEdges {
                                top: floating.top_from_text.unwrap_or(0.0),
                                right: floating.right_from_text.unwrap_or(0.0),
                                bottom: floating.bottom_from_text.unwrap_or(0.0),
                                left: floating.left_from_text.unwrap_or(0.0),
                            },
                            anchor_y,
                            false,
                            false,
                            None,
                        );
                    }
                }
            }
            _ => {}
        }
    }
    bands.sort_by(|a, b| a.top.total_cmp(&b.top));
    bands
}

pub fn extend_body_margins(
    page_size: &Size,
    margins: &PageMargins,
    header_height: f64,
    footer_height: f64,
) -> PageMargins {
    let header_distance = margins.header.unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let footer_distance = margins.footer.unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let suppress_header = margins.top < 0.0;
    let suppress_footer = margins.bottom < 0.0;
    let effective_top = margins.top.abs();
    let effective_bottom = margins.bottom.abs();
    let mut output = margins.clone();
    output.top = effective_top;
    output.bottom = effective_bottom;
    if !suppress_header && header_height > effective_top - header_distance {
        output.top = effective_top.max(header_distance + header_height);
    }
    if !suppress_footer && footer_height > effective_bottom - footer_distance {
        output.bottom = effective_bottom.max(footer_distance + footer_height);
    }
    let maximum = (page_size.h - MIN_CONTENT_HEIGHT_PX).max(0.0);
    if output.top + output.bottom > maximum {
        output.bottom = output.bottom.min((maximum - output.top).max(0.0));
        if output.top + output.bottom > maximum {
            output.top = (maximum - output.bottom).max(0.0);
        }
    }
    output
}

#[cfg(test)]
mod tests {
    use crate::display_list::{DisplayList, Primitive, build_display_list_json};
    use serde_json::json;

    use super::*;

    fn header_with_image(
        relative: &str,
        offset: f64,
        trailing_text: bool,
    ) -> (HeaderFooterVariant, Size, PageMargins) {
        let mut blocks = vec![json!({
            "kind": "paragraph", "id": "anchor",
            "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
            "runs": [
                {"kind": "text", "text": "Anchor"},
                {"kind": "image", "src": "image", "width": 100, "height": 100,
                 "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                 "position": {"vertical": {"relativeTo": relative, "posOffset": offset * 9525.0}}}
            ]
        })];
        if trailing_text {
            blocks.push(json!({
                "kind": "paragraph", "id": "tail",
                "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                "runs": [{"kind": "text", "text": "Tail"}]
            }));
        }
        header_footer_with_blocks(HeaderFooterKind::Header, blocks)
    }

    fn header_footer_measurement_config() -> MeasurementConfig {
        let font_id = crate::register_measure_font(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        serde_json::from_value(json!({
            "fontChains": {"liberation sans|0|0": [font_id]},
            "defaults": {"fontFamily": "Liberation Sans", "fontSize": 12}
        }))
        .unwrap()
    }

    fn header_footer_with_blocks(
        kind: HeaderFooterKind,
        blocks: Vec<serde_json::Value>,
    ) -> (HeaderFooterVariant, Size, PageMargins) {
        let config = header_footer_measurement_config();
        let size = Size { w: 500.0, h: 500.0 };
        let margins = PageMargins {
            top: 96.0,
            right: 96.0,
            bottom: 96.0,
            left: 96.0,
            header: Some(48.0),
            footer: Some(48.0),
        };
        let variant = measure_header_footer(
            "hf".to_owned(),
            kind,
            HeaderFooterType::Default,
            0,
            serde_json::from_value(json!(blocks)).unwrap(),
            308.0,
            HeaderFooterMetrics {
                kind,
                page_size: &size,
                margins: &margins,
            },
            &config,
        )
        .unwrap()
        .unwrap();
        (variant, size, margins)
    }

    #[test]
    fn in_front_header_image_preserves_unrounded_visual_top() {
        let (variant, size, margins) = header_footer_with_blocks(
            HeaderFooterKind::Header,
            vec![json!({
                "kind": "paragraph", "id": "anchor",
                "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                "runs": [
                    {"kind": "text", "text": "Anchor"},
                    {"kind": "image", "src": "image", "width": 100, "height": 100,
                     "wrapType": "inFront",
                     "position": {"vertical": {"relativeTo": "page", "posOffset": 453390}}}
                ]
            })],
        );
        assert!((variant.visual_top + 0.4).abs() < 1e-9);
        assert_eq!(variant.visual_bottom, 103.390625);
        assert_eq!(variant.flow_height, 103.390625);
        let serialized = serde_json::to_value(&variant).unwrap();
        assert!((serialized["visualTop"].as_f64().unwrap() + 0.4).abs() < 1e-9);
        assert!(
            header_footer_float_bands(
                &variant,
                HeaderFooterMetrics {
                    kind: HeaderFooterKind::Header,
                    page_size: &size,
                    margins: &margins,
                },
            )
            .is_empty()
        );
    }

    #[test]
    fn behind_document_top_and_bottom_header_image_keeps_original_flow_and_body_top() {
        let blocks = vec![json!({
            "kind": "paragraph", "id": "anchor",
            "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
            "runs": [
                {"kind": "text", "text": "Anchor"},
                {"kind": "image", "src": "image", "width": 100, "height": 100,
                 "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                 "position": {"behindDoc": true,
                              "vertical": {"relativeTo": "page", "posOffset": 200 * 9525}}}
            ]
        })];
        let mut original: Vec<LayoutBlock> = serde_json::from_value(json!(blocks)).unwrap();
        let measures =
            measure_blocks(&mut original, 308.0, &header_footer_measurement_config()).unwrap();
        let expected_height = measures.iter().map(extent_height).sum::<f64>();
        let expected: Vec<_> = original
            .into_iter()
            .zip(measures)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect();
        let (variant, size, margins) = header_footer_with_blocks(HeaderFooterKind::Header, blocks);
        assert_eq!(
            serde_json::to_value(&variant.measured).unwrap(),
            serde_json::to_value(&expected).unwrap()
        );
        assert_eq!(variant.flow_height, expected_height);
        assert!(variant.flow_height > 100.0);
        assert!(
            header_footer_float_bands(
                &variant,
                HeaderFooterMetrics {
                    kind: HeaderFooterKind::Header,
                    page_size: &size,
                    margins: &margins,
                },
            )
            .is_empty()
        );
        let body_margins = extend_body_margins(&size, &margins, variant.flow_height, 0.0);
        assert_eq!(body_margins.top, margins.header.unwrap() + expected_height);
        let mut input: crate::types::Input = serde_json::from_value(json!({
            "measured": [{
                "block": {"kind": "paragraph", "id": "body",
                          "runs": [{"kind": "text", "text": "Body"}]},
                "measure": {"kind": "paragraph", "totalHeight": 20,
                            "lines": [{"headRun": 0, "headChar": 0, "tailRun": 0,
                                       "tailChar": 4, "width": 40, "ascent": 15,
                                       "descent": 5, "lineHeight": 20}]}
            }],
            "options": {"pageSize": size, "margins": body_margins}
        }))
        .unwrap();
        let layout = crate::place::layout_document(&mut input).unwrap();
        let crate::types::Fragment::Paragraph(body) = &layout.pages[0].fragments[0] else {
            panic!("paragraph expected");
        };
        assert_eq!(body.y, margins.header.unwrap() + expected_height);
    }

    #[test]
    fn another_sections_header_float_preserves_effective_body_anchor_margins() {
        let (mut header, size, margins) = header_footer_with_blocks(
            HeaderFooterKind::Header,
            vec![json!({
                "kind": "paragraph", "id": "tall-header",
                "attrs": {"spacing": {"line": 152, "lineRule": "exact"}},
                "runs": [{"kind": "text", "text": "Header"}],
            })],
        );
        header.r_id = "tall-header".to_owned();
        assert_eq!(header.flow_height, 152.0);
        let effective = extend_body_margins(&size, &margins, header.flow_height, 0.0);
        assert_eq!(effective.top, 200.0);
        let (mut float_header, _, _) = header_with_image("margin", 0.0, false);
        float_header.r_id = "float-header".to_owned();
        float_header.section_index = 1;
        let bands = header_footer_float_bands(
            &float_header,
            HeaderFooterMetrics {
                kind: HeaderFooterKind::Header,
                page_size: &size,
                margins: &margins,
            },
        );
        let position = json!({
            "horizontal": {"relativeTo": "margin", "posOffset": 0},
            "vertical": {"relativeTo": "margin", "posOffset": 0},
        });
        let mut input: crate::types::Input = serde_json::from_value(json!({
            "measured": [
                {
                    "block": {"kind": "paragraph", "id": "body", "runs": [
                        {"kind": "text", "text": "Body"},
                        {"kind": "image", "src": "body-image", "width": 20, "height": 20,
                         "displayMode": "float", "wrapType": "square", "position": position},
                    ]},
                    "measure": {"kind": "paragraph", "totalHeight": 20, "lines": [{
                        "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                        "width": 40, "ascent": 15, "descent": 5, "lineHeight": 20,
                    }]},
                },
                {
                    "block": {"kind": "textBox", "id": "body-box", "width": 20,
                              "height": 20, "displayMode": "float", "position": position,
                              "fillColor": "#eeeeee", "content": []},
                    "measure": {"kind": "textBox", "width": 20, "height": 20,
                                "innerMeasures": []},
                },
                {"block": {"kind": "sectionBreak", "id": "break", "type": "nextPage",
                           "margins": effective}, "measure": {"kind": "sectionBreak"}},
                {
                    "block": {"kind": "paragraph", "id": "tail",
                              "runs": [{"kind": "text", "text": "Tail"}]},
                    "measure": {"kind": "paragraph", "totalHeight": 20, "lines": [{
                        "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                        "width": 40, "ascent": 15, "descent": 5, "lineHeight": 20,
                    }]},
                },
            ],
            "options": {"pageSize": size, "margins": effective, "finalMargins": margins,
                        "bodyBreakType": "nextPage", "sectionPageFloatBands": [
                            {"default": [], "anchorMargins": margins},
                            {"default": bands, "anchorMargins": margins},
                        ]},
        }))
        .unwrap();
        let regions = serde_json::from_value(json!({"sections": [
            {"headerFooterRefs": {"headerDefault": header.r_id}},
            {"headerFooterRefs": {"headerDefault": float_header.r_id}},
        ]}))
        .unwrap();
        let mut layout = crate::place::layout_document(&mut input).unwrap();
        crate::regions::apply_document_regions(&mut layout, &regions);
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(layout.pages[0].margins, effective);
        assert!(layout.pages[0].body_margins.is_none());
        assert!(layout.pages[0].body_anchor_margins.is_none());
        let body_box = layout.pages[0]
            .fragments
            .iter()
            .find_map(|fragment| match fragment {
                crate::types::Fragment::TextBox(fragment) => Some(fragment),
                _ => None,
            })
            .unwrap();
        assert_eq!((body_box.x, body_box.y), (96.0, 200.0));
        let display: DisplayList = serde_json::from_str(
            &build_display_list_json(
                &json!({"measured": input.measured, "options": input.options, "layout": layout,
                        "headersFooters": {"variants": [header, float_header]}})
                .to_string(),
            )
            .unwrap(),
        )
        .unwrap();
        let body_image = display.pages[0]
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Image(image) => Some(image),
                _ => None,
            })
            .unwrap();
        assert_eq!(body_image.y.as_f64(), Some(200.0));
        let header_image = display.pages[1]
            .header
            .as_ref()
            .unwrap()
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Image(image) => Some(image),
                _ => None,
            })
            .unwrap();
        assert_eq!(header_image.y.as_f64(), Some(96.0));
    }

    #[test]
    fn folded_body_margins_preserve_the_header_image_anchor() {
        let (variant, size, margins) = header_with_image("margin", 0.0, false);
        let regions = serde_json::from_value(json!({
            "sections": [{"headerFooterRefs": {"headerDefault": variant.r_id}}]
        }))
        .unwrap();
        let bands = header_footer_float_bands(
            &variant,
            HeaderFooterMetrics {
                kind: HeaderFooterKind::Header,
                page_size: &size,
                margins: &margins,
            },
        );
        for body_top in [96.0_f64, 224.0] {
            let mut input: crate::types::Input = serde_json::from_value(json!({
                "measured": [{
                    "block": {"kind": "paragraph", "id": "body",
                              "runs": [
                                  {"kind": "text", "text": "Body"},
                                  {"kind": "image", "src": "body-image", "width": 100, "height": 100,
                                   "displayMode": "float", "wrapType": "topAndBottom",
                                   "position": {"vertical": {"relativeTo": "margin", "posOffset": 0}}}
                              ]},
                    "measure": {"kind": "paragraph", "totalHeight": 20,
                                "lines": [{"headRun": 0, "headChar": 0, "tailRun": 0,
                                           "tailChar": 4, "width": 40, "ascent": 15,
                                           "descent": 5, "lineHeight": 20}]}
                }],
                "options": {"pageSize": size, "margins": margins,
                            "sectionPageFloatBands": [{"default": bands, "anchorMargins": margins}]}
            }))
            .unwrap();
            input.options.margins.as_mut().unwrap().top = body_top;
            let expected_top = body_top.max(196.0);
            let mut layout = crate::place::layout_document(&mut input).unwrap();
            crate::regions::apply_document_regions(&mut layout, &regions);
            let page = &layout.pages[0];
            assert_eq!(page.margins.top, 96.0);
            assert_eq!(page.body_margins.as_ref().unwrap().top, expected_top);
            let crate::types::Fragment::Paragraph(body) = &page.fragments[0] else {
                panic!("paragraph expected");
            };
            assert!(body.y >= expected_top);
            let display: DisplayList = serde_json::from_str(
                &build_display_list_json(
                    &json!({
                        "measured": input.measured, "options": input.options, "layout": layout,
                        "headersFooters": {"variants": [variant]}
                    })
                    .to_string(),
                )
                .unwrap(),
            )
            .unwrap();
            let image = display.pages[0]
                .header
                .as_ref()
                .unwrap()
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    Primitive::Image(image) => Some(image),
                    _ => None,
                })
                .unwrap();
            assert_eq!(image.y.as_f64(), Some(96.0));
            assert_eq!(image.y.as_f64().unwrap() + image.h.as_f64().unwrap(), 196.0);
            assert_eq!(
                display.pages[0].content_bounds.as_ref().unwrap().y.as_f64(),
                Some(expected_top)
            );
            let body_image = display.pages[0]
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    Primitive::Image(image) => Some(image),
                    _ => None,
                })
                .unwrap();
            assert_eq!(body_image.y.as_f64(), Some(body_top));
        }
    }

    #[test]
    fn detached_header_image_keeps_later_text_below_it() {
        let (variant, _, margins) = header_with_image("page", 64.0, true);
        let LayoutBlock::Paragraph(anchor) = &variant.measured[0].block else {
            panic!("paragraph expected");
        };
        let Run::Image(image) = &anchor.runs[1] else {
            panic!("image expected");
        };
        assert_ne!(image.display_mode.as_deref(), Some("float"));
        let mut flow = HeaderFooterFlow::default();
        let mut tail_top = 0.0;
        for measured in &variant.measured {
            let (before, after) = block_spacing(&measured.block);
            tail_top = flow.place(
                (extent_height(&measured.measure) - before - after).max(0.0),
                before,
                after,
            );
        }
        assert!(margins.header.unwrap() + tail_top >= 164.0);
    }

    #[test]
    fn detached_header_image_stays_floating_when_later_text_is_clear() {
        let (variant, _, _) = header_with_image("page", 200.0, true);
        let LayoutBlock::Paragraph(anchor) = &variant.measured[0].block else {
            panic!("paragraph expected");
        };
        let Run::Image(image) = &anchor.runs[1] else {
            panic!("image expected");
        };
        assert_eq!(image.display_mode.as_deref(), Some("float"));
        assert_eq!(variant.flow_height, 32.0);
    }

    #[test]
    fn seventeen_detached_images_keep_original_measurement() {
        let blocks: Vec<_> = (0..17)
            .map(|index| {
                json!({
                    "kind": "paragraph", "id": index,
                    "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                    "runs": [
                        {"kind": "text", "text": "Anchor"},
                        {"kind": "image", "src": "image", "width": 20, "height": 20,
                         "displayMode": if index % 2 == 0 { None } else { Some("block") },
                         "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                         "position": {"vertical": {"relativeTo": "page", "posOffset": 400 * 9525}}}
                    ]
                })
            })
            .collect();
        let (variant, _, _) = header_footer_with_blocks(HeaderFooterKind::Header, blocks.clone());
        let mut original: Vec<LayoutBlock> = serde_json::from_value(json!(blocks)).unwrap();
        apply_contextual_spacing_blocks(&mut original);
        let measures =
            measure_blocks(&mut original, 308.0, &header_footer_measurement_config()).unwrap();
        let expected: Vec<_> = original
            .into_iter()
            .zip(measures)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect();
        assert_eq!(
            serde_json::to_value(&variant.measured).unwrap(),
            serde_json::to_value(&expected).unwrap()
        );
        assert_eq!(
            variant.flow_height,
            expected
                .iter()
                .map(|measured| extent_height(&measured.measure))
                .sum::<f64>()
        );
    }

    #[test]
    fn overlapping_detached_image_restores_all_images() {
        let (variant, _, _) = header_footer_with_blocks(
            HeaderFooterKind::Header,
            vec![json!({
                "kind": "paragraph", "id": "anchor",
                "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                "runs": [
                    {"kind": "text", "text": "Anchor"},
                    {"kind": "image", "src": "overlapping", "width": 100, "height": 100,
                     "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                     "position": {"vertical": {"relativeTo": "page", "posOffset": 48 * 9525}}},
                    {"kind": "image", "src": "clear", "width": 20, "height": 20,
                     "displayMode": "block", "wrapType": "topAndBottom",
                     "distTop": 0, "distBottom": 0,
                     "position": {"vertical": {"relativeTo": "page", "posOffset": 400 * 9525}}}
                ]
            })],
        );
        let LayoutBlock::Paragraph(paragraph) = &variant.measured[0].block else {
            panic!("paragraph expected");
        };
        let Run::Image(overlapping) = &paragraph.runs[1] else {
            panic!("image expected");
        };
        let Run::Image(clear) = &paragraph.runs[2] else {
            panic!("image expected");
        };
        assert_eq!(overlapping.display_mode, None);
        assert_eq!(clear.display_mode.as_deref(), Some("block"));
    }

    #[test]
    fn detached_header_image_keeps_same_paragraph_tail_below_it() {
        let (variant, _, margins) = header_footer_with_blocks(
            HeaderFooterKind::Header,
            vec![json!({
                "kind": "paragraph", "id": "anchor",
                "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                "runs": [
                    {"kind": "image", "src": "image", "width": 100, "height": 100,
                     "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                     "position": {"vertical": {"relativeTo": "page", "posOffset": 48 * 9525}}},
                    {"kind": "text", "text": "Tail"}
                ]
            })],
        );
        let LayoutBlock::Paragraph(paragraph) = &variant.measured[0].block else {
            panic!("paragraph expected");
        };
        let Run::Image(image) = &paragraph.runs[0] else {
            panic!("image expected");
        };
        assert_ne!(image.display_mode.as_deref(), Some("float"));
        let BlockExtent::Paragraph(measure) = &variant.measured[0].measure else {
            panic!("paragraph expected");
        };
        let tail_line = measure
            .lines
            .iter()
            .position(|line| line.head_run == 1)
            .unwrap();
        let tail_top = margins.header.unwrap()
            + measure.lines[..tail_line]
                .iter()
                .map(|line| line.float_skip_before.unwrap_or(0.0) + line.line_height)
                .sum::<f64>();
        assert!(tail_top >= 148.0);
    }

    #[test]
    fn detached_footer_image_keeps_preceding_text_clear() {
        let (variant, size, margins) = header_footer_with_blocks(
            HeaderFooterKind::Footer,
            vec![
                json!({
                    "kind": "paragraph", "id": "text",
                    "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                    "runs": [{"kind": "text", "text": "Tail"}]
                }),
                json!({
                    "kind": "paragraph", "id": "anchor",
                    "attrs": {"spacing": {"line": 16, "lineRule": "exact"}},
                    "runs": [
                        {"kind": "image", "src": "image", "width": 100, "height": 100,
                         "wrapType": "topAndBottom", "distTop": 0, "distBottom": 0,
                         "position": {"vertical": {"relativeTo": "page", "posOffset": 400 * 9525}}}
                    ]
                }),
            ],
        );
        let LayoutBlock::Paragraph(paragraph) = &variant.measured[1].block else {
            panic!("paragraph expected");
        };
        let Run::Image(image) = &paragraph.runs[0] else {
            panic!("image expected");
        };
        assert_ne!(image.display_mode.as_deref(), Some("float"));
        let text_top = size.h - margins.footer.unwrap() - variant.flow_height;
        assert!(text_top + extent_height(&variant.measured[0].measure) <= 400.0);
    }

    #[test]
    fn inside_and_outside_margin_bands_match_the_painted_header_image() {
        for relative in ["insideMargin", "outsideMargin"] {
            let (variant, size, margins) = header_with_image(relative, 0.0, false);
            let bands = header_footer_float_bands(
                &variant,
                HeaderFooterMetrics {
                    kind: HeaderFooterKind::Header,
                    page_size: &size,
                    margins: &margins,
                },
            );
            assert_eq!(bands.len(), 2);
            assert!(
                bands
                    .iter()
                    .all(|band| band.top == 96.0 && band.bottom == 196.0)
            );
            let mut input: crate::types::Input = serde_json::from_value(json!({
                "measured": [{
                    "block": {"kind": "paragraph", "id": "body",
                              "runs": [{"kind": "text", "text": "Body"}]},
                    "measure": {"kind": "paragraph", "totalHeight": 20,
                                "lines": [{"headRun": 0, "headChar": 0, "tailRun": 0,
                                           "tailChar": 4, "width": 40, "ascent": 15,
                                           "descent": 5, "lineHeight": 20}]}
                }],
                "options": {"pageSize": size, "margins": margins,
                            "sectionPageFloatBands": [{"default": bands, "anchorMargins": margins}]}
            }))
            .unwrap();
            let mut layout = crate::place::layout_document(&mut input).unwrap();
            let regions = serde_json::from_value(json!({
                "sections": [{"headerFooterRefs": {"headerDefault": variant.r_id}}]
            }))
            .unwrap();
            crate::regions::apply_document_regions(&mut layout, &regions);
            let crate::types::Fragment::Paragraph(body) = &layout.pages[0].fragments[0] else {
                panic!("paragraph expected");
            };
            assert!(body.y >= 196.0);
            let display: DisplayList = serde_json::from_str(
                &build_display_list_json(
                    &json!({
                        "measured": input.measured, "options": input.options, "layout": layout,
                        "headersFooters": {"variants": [variant]}
                    })
                    .to_string(),
                )
                .unwrap(),
            )
            .unwrap();
            let image = display.pages[0]
                .header
                .as_ref()
                .unwrap()
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    Primitive::Image(image) => Some(image),
                    _ => None,
                })
                .unwrap();
            assert_eq!(image.y.as_f64(), Some(96.0));
            assert_eq!(image.y.as_f64().unwrap() + image.h.as_f64().unwrap(), 196.0);
        }
    }

    #[test]
    fn empty_paragraph_after_table_reserves_header_footer_space() {
        for kind in [HeaderFooterKind::Header, HeaderFooterKind::Footer] {
            let blocks = serde_json::from_value(json!([
                {"kind":"table","id":"table","rows":[{"id":"row","height":20,"heightRule":"exact","cells":[{"id":"cell","blocks":[]}]}]},
                {"kind":"paragraph","id":"tail","runs":[],"attrs":{"spacing":{"before":2,"after":3,"line":12,"lineRule":"exact"}}}
            ])).unwrap();
            let size = Size { w: 300.0, h: 500.0 };
            let margins = PageMargins {
                top: 40.0,
                right: 40.0,
                bottom: 40.0,
                left: 40.0,
                header: Some(20.0),
                footer: Some(20.0),
            };
            let variant = measure_header_footer(
                "hf".to_owned(),
                kind,
                HeaderFooterType::Default,
                0,
                blocks,
                220.0,
                HeaderFooterMetrics {
                    kind,
                    page_size: &size,
                    margins: &margins,
                },
                &MeasurementConfig::default(),
            )
            .unwrap()
            .unwrap();
            assert_eq!(variant.flow_height, 37.0);
            let BlockExtent::Paragraph(tail) = &variant.measured[1].measure else {
                panic!("paragraph expected");
            };
            assert_eq!(tail.total_height, 17.0);
            assert_eq!(tail.lines[0].line_height, 12.0);
        }
    }

    #[test]
    fn measured_header_height_includes_collapsed_style_spacing() {
        let blocks = serde_json::from_value(json!([
            {"kind":"paragraph","id":"a","runs":[{"kind":"text","text":"A"}],"attrs":{"spacing":{"before":5,"after":8}}},
            {"kind":"paragraph","id":"b","runs":[{"kind":"text","text":"B"}],"attrs":{"spacing":{"before":4,"after":6}}}
        ])).unwrap();
        let size = Size { w: 300.0, h: 500.0 };
        let margins = PageMargins {
            top: 40.0,
            right: 40.0,
            bottom: 40.0,
            left: 40.0,
            header: Some(20.0),
            footer: Some(20.0),
        };
        let variant = measure_header_footer(
            "header".to_owned(),
            HeaderFooterKind::Header,
            HeaderFooterType::Default,
            0,
            blocks,
            220.0,
            HeaderFooterMetrics {
                kind: HeaderFooterKind::Header,
                page_size: &size,
                margins: &margins,
            },
            &MeasurementConfig::default(),
        )
        .unwrap()
        .unwrap();
        let text_height: f64 = variant
            .measured
            .iter()
            .map(|m| match &m.measure {
                BlockExtent::Paragraph(p) => {
                    p.lines.iter().map(|line| line.line_height).sum::<f64>()
                }
                _ => panic!("paragraph expected"),
            })
            .sum();
        assert_eq!(variant.flow_height, text_height + 5.0 + 8.0 + 6.0);
        assert_eq!(variant.visual_bottom, variant.flow_height);
    }

    #[test]
    fn margin_extension_uses_flow_height_and_preserves_body_floor() {
        let margins = PageMargins {
            top: 96.0,
            right: 96.0,
            bottom: 96.0,
            left: 96.0,
            header: Some(48.0),
            footer: Some(48.0),
        };
        let page_size = Size { w: 816.0, h: 200.0 };

        let extended = extend_body_margins(&page_size, &margins, 140.0, 100.0);
        assert_eq!(extended.top + extended.bottom, 176.0);
        assert_eq!(extended.bottom, 0.0);
    }

    #[test]
    fn negative_top_uses_absolute_origin_and_ignores_header() {
        let margins = PageMargins {
            top: -1438.0 / 15.0,
            right: 1797.0 / 15.0,
            bottom: 96.0,
            left: 1797.0 / 15.0,
            header: Some(709.0 / 15.0),
            footer: Some(48.0),
        };
        let page_size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let extended = extend_body_margins(&page_size, &margins, 100.0, 0.0);
        assert_eq!(extended.top, 1438.0 / 15.0);
        assert_eq!(extended.bottom, 96.0);
    }

    #[test]
    fn negative_bottom_uses_absolute_origin_and_ignores_footer() {
        let margins = PageMargins {
            top: 96.0,
            right: 96.0,
            bottom: -1440.0 / 15.0,
            left: 96.0,
            header: Some(48.0),
            footer: Some(48.0),
        };
        let page_size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let extended = extend_body_margins(&page_size, &margins, 0.0, 100.0);
        assert_eq!(extended.top, 96.0);
        assert_eq!(extended.bottom, 1440.0 / 15.0);
    }

    #[test]
    fn both_negative_use_absolute_origins_without_expansion() {
        let margins = PageMargins {
            top: -1438.0 / 15.0,
            right: 96.0,
            bottom: -1440.0 / 15.0,
            left: 96.0,
            header: Some(709.0 / 15.0),
            footer: Some(709.0 / 15.0),
        };
        let page_size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let extended = extend_body_margins(&page_size, &margins, 100.0, 100.0);
        assert_eq!(extended.top, 1438.0 / 15.0);
        assert_eq!(extended.bottom, 1440.0 / 15.0);
    }

    #[test]
    fn positive_margins_expand_for_header_overflow() {
        let margins = PageMargins {
            top: 40.0,
            right: 40.0,
            bottom: 40.0,
            left: 40.0,
            header: Some(20.0),
            footer: Some(20.0),
        };
        let page_size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let extended = extend_body_margins(&page_size, &margins, 50.0, 0.0);
        assert_eq!(extended.top, 70.0);
        assert_eq!(extended.bottom, 40.0);
    }

    #[test]
    fn negative_margins_without_headers_keep_absolute_origin() {
        let margins = PageMargins {
            top: -60.0,
            right: 96.0,
            bottom: -70.0,
            left: 96.0,
            header: None,
            footer: None,
        };
        let page_size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let extended = extend_body_margins(&page_size, &margins, 0.0, 0.0);
        assert_eq!(extended.top, 60.0);
        assert_eq!(extended.bottom, 70.0);
    }

    #[test]
    fn negative_margins_respect_page_capacity() {
        let margins = PageMargins {
            top: -140.0,
            right: 96.0,
            bottom: 100.0,
            left: 96.0,
            header: Some(48.0),
            footer: Some(48.0),
        };
        let page_size = Size { w: 816.0, h: 200.0 };
        let extended = extend_body_margins(&page_size, &margins, 0.0, 0.0);
        assert_eq!(extended.top, 140.0);
        assert_eq!(extended.bottom, 36.0);
    }

    #[test]
    fn page_field_widths_resolve_from_final_page_labels() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        crate::clear_measure_fonts();
        let font_id = crate::register_measure_font(FONT).unwrap();
        let config: MeasurementConfig = serde_json::from_value(json!({
            "fontChains": {"liberation sans|0|0": [font_id]},
            "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"}
        }))
        .unwrap();
        let measured: MeasuredBlock = serde_json::from_value(json!({
            "block": {
                "kind": "paragraph",
                "id": "field-paragraph",
                "runs": [{
                    "kind": "field",
                    "fieldType": "PAGE",
                    "fallback": "1",
                    "fontFamily": "Liberation Sans",
                    "fontSize": 11,
                    "pmStart": 2,
                    "pmEnd": 3
                }]
            },
            "measure": {"kind": "paragraph", "lines": [], "totalHeight": 0}
        }))
        .unwrap();
        let mut input: crate::types::Input = serde_json::from_value(json!({
            "measured": [],
            "options": {}
        }))
        .unwrap();
        let mut layout = crate::place::layout_document(&mut input).unwrap();
        let mut second = layout.pages[0].clone();
        second.number = 2;
        second.page_label = Some("VIII".to_owned());
        layout.pages.push(second);
        let mut payload = HeaderFooterPayload {
            variants: vec![HeaderFooterVariant {
                r_id: "rId1".to_owned(),
                kind: HeaderFooterKind::Footer,
                hf_type: HeaderFooterType::Default,
                section_index: 0,
                measured: vec![measured],
                height: 0.0,
                flow_height: 0.0,
                visual_top: 0.0,
                visual_bottom: 0.0,
                field_widths: Vec::new(),
            }],
            ..HeaderFooterPayload::default()
        };

        resolve_header_footer_field_widths(&mut payload, &layout, &config).unwrap();

        let widths = &payload.variants[0].field_widths[0];
        assert_eq!(widths.pm_start, 2);
        assert_eq!(widths.fallback_width, widths.per_page[0]);
        assert!(widths.per_page[1] > widths.per_page[0]);
    }

    #[test]
    fn numpages_of_a_partial_layout_takes_no_width() {
        const FONT: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        crate::clear_measure_fonts();
        let font_id = crate::register_measure_font(FONT).unwrap();
        let config: MeasurementConfig = serde_json::from_value(json!({
            "fontChains": {"liberation sans|0|0": [font_id]},
            "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"}
        }))
        .unwrap();
        let measured: MeasuredBlock = serde_json::from_value(json!({
            "block": {
                "kind": "paragraph",
                "id": "field-paragraph",
                "runs": [{
                    "kind": "field",
                    "fieldType": "NUMPAGES",
                    "fallback": "9",
                    "fontFamily": "Liberation Sans",
                    "fontSize": 11,
                    "pmStart": 2,
                    "pmEnd": 3
                }]
            },
            "measure": {"kind": "paragraph", "lines": [], "totalHeight": 0}
        }))
        .unwrap();
        let mut input: crate::types::Input = serde_json::from_value(json!({
            "measured": [],
            "options": {}
        }))
        .unwrap();
        let mut layout = crate::place::layout_document(&mut input).unwrap();
        let payload = || HeaderFooterPayload {
            variants: vec![HeaderFooterVariant {
                r_id: "rId1".to_owned(),
                kind: HeaderFooterKind::Footer,
                hf_type: HeaderFooterType::Default,
                section_index: 0,
                measured: vec![measured.clone()],
                height: 0.0,
                flow_height: 0.0,
                visual_top: 0.0,
                visual_bottom: 0.0,
                field_widths: Vec::new(),
            }],
            ..HeaderFooterPayload::default()
        };

        let mut whole = payload();
        resolve_header_footer_field_widths(&mut whole, &layout, &config).unwrap();
        assert!(whole.variants[0].field_widths[0].per_page[0] > 0.0);

        layout.partial = true;
        let mut partial = payload();
        resolve_header_footer_field_widths(&mut partial, &layout, &config).unwrap();
        assert_eq!(partial.variants[0].field_widths[0].per_page[0], 0.0);

        layout.cached_page_totals = true;
        let mut cached = payload();
        resolve_header_footer_field_widths(&mut cached, &layout, &config).unwrap();
        let widths = &cached.variants[0].field_widths[0];
        assert!(widths.per_page[0] > 0.0);
        assert_eq!(widths.per_page[0], widths.fallback_width);

        for fallback in [None, Some(String::new())] {
            if let LayoutBlock::Paragraph(paragraph) = &mut cached.variants[0].measured[0].block {
                let Run::Field(field) = &mut paragraph.runs[0] else {
                    panic!("expected a field");
                };
                field.fallback = fallback;
            }
            resolve_header_footer_field_widths(&mut cached, &layout, &config).unwrap();
            assert_eq!(cached.variants[0].field_widths[0].per_page[0], 0.0);
        }
    }
}
