use std::borrow::Cow;

use serde::Serialize;

use crate::measure_blocks::{MeasurementConfig, extent_height, measure_blocks, measure_paragraph};
use crate::types::{
    BlockExtent, BlockId, FieldRun, ImageRun, Layout, LayoutBlock, MeasuredBlock, PageMargins,
    ParagraphBlock, Run, RunFormatting, Size, TableBlock,
};

const DEFAULT_HF_DISTANCE_PX: f64 = 48.0;
const MIN_CONTENT_HEIGHT_PX: f64 = 24.0;

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
    let mut blocks = normalize_header_footer_blocks(blocks);
    let measures = measure_blocks(&mut blocks, content_width, config)?;
    let height = measures.iter().map(extent_height).sum();
    let flow_height = blocks
        .iter()
        .zip(&measures)
        .filter(|(block, _)| contributes_to_flow(block))
        .map(|(_, measure)| extent_height(measure))
        .sum();
    let (visual_top, visual_bottom) = visual_bounds(&blocks, &measures, height, metrics);
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

pub fn resolve_header_footer_field_widths(
    payload: &mut HeaderFooterPayload,
    layout: &Layout,
    config: &MeasurementConfig,
) -> Result<(), String> {
    let total_pages = layout.pages.len().to_string();
    for variant in &mut payload.variants {
        let mut widths = Vec::new();
        for measured in &variant.measured {
            let LayoutBlock::Paragraph(paragraph) = &measured.block else {
                continue;
            };
            for run in &paragraph.runs {
                let Run::Field(field) = run else {
                    continue;
                };
                if !matches!(field.field_type.as_str(), "PAGE" | "NUMPAGES") {
                    continue;
                }
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
                            page.page_label
                                .as_deref()
                                .map(Cow::Borrowed)
                                .unwrap_or_else(|| Cow::Owned(page.number.to_string()))
                        };
                        measure_field_text(field, &text, config)
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
            fmt: RunFormatting {
                bold: field.fmt.bold,
                italic: field.fmt.italic,
                font_family: field.fmt.font_family.clone(),
                font_size: field.fmt.font_size,
                ..RunFormatting::default()
            },
            field_type: field.field_type.clone(),
            raw_type: None,
            instruction: None,
            fallback: Some(text.to_owned()),
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

pub fn normalize_header_footer_blocks(mut blocks: Vec<LayoutBlock>) -> Vec<LayoutBlock> {
    normalize_block_slice(&mut blocks);
    blocks
}

fn normalize_block_slice(blocks: &mut [LayoutBlock]) {
    let trailing_empty: Vec<usize> = (1..blocks.len())
        .filter(|index| {
            matches!(blocks[index - 1], LayoutBlock::Table(_))
                && matches!(&blocks[*index], LayoutBlock::Paragraph(paragraph)
                    if paragraph.runs.is_empty() && !has_authored_visuals(paragraph))
        })
        .collect();
    for block in blocks.iter_mut() {
        if let LayoutBlock::Table(table) = block {
            normalize_table(table);
        }
        let LayoutBlock::Paragraph(paragraph) = block else {
            continue;
        };
        let Some(attrs) = paragraph.attrs.as_mut() else {
            continue;
        };
        if let Some(spacing) = attrs.spacing.as_mut() {
            let explicit = attrs.spacing_explicit.as_ref();
            if explicit.and_then(|value| value.before) != Some(true) {
                spacing.before = None;
            }
            if explicit.and_then(|value| value.after) != Some(true) {
                spacing.after = None;
            }
        }
    }
    for index in trailing_empty {
        let LayoutBlock::Paragraph(paragraph) = &mut blocks[index] else {
            continue;
        };
        paragraph
            .attrs
            .get_or_insert_with(Default::default)
            .suppress_empty_paragraph_height = Some(true);
    }
}

fn normalize_table(table: &mut TableBlock) {
    for row in &mut table.rows {
        for cell in &mut row.cells {
            normalize_block_slice(&mut cell.blocks);
        }
    }
}

fn has_authored_visuals(paragraph: &crate::types::ParagraphBlock) -> bool {
    let Some(attrs) = &paragraph.attrs else {
        return false;
    };
    attrs
        .borders
        .as_ref()
        .is_some_and(|borders| borders.top.is_some() || borders.bottom.is_some())
        || attrs
            .spacing_explicit
            .as_ref()
            .is_some_and(|spacing| spacing.before == Some(true) || spacing.after == Some(true))
}

pub fn contributes_to_flow(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(_) | LayoutBlock::Table(_) => true,
        LayoutBlock::Image(image) => {
            image.anchor.as_ref().and_then(|anchor| anchor.is_anchored) != Some(true)
        }
        LayoutBlock::Shape(_) | LayoutBlock::Chart(_) => true,
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
    let mut cursor = 0.0_f64;
    for (block, measure) in blocks.iter().zip(measures) {
        let block_height = extent_height(measure);
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
                    let top = image_visual_top(image, cursor, height, metrics);
                    visual_top = visual_top.min(top);
                    visual_bottom = visual_bottom.max(top + image.height);
                }
                cursor += block_height;
            }
            LayoutBlock::TextBox(text_box) => {
                visual_top = visual_top.min(cursor);
                visual_bottom = visual_bottom.max(cursor + block_height);
                if text_box.display_mode.as_deref() != Some("float") {
                    cursor += block_height;
                }
            }
            LayoutBlock::Table(_)
            | LayoutBlock::Image(_)
            | LayoutBlock::Shape(_)
            | LayoutBlock::Chart(_) => {
                visual_top = visual_top.min(cursor);
                visual_bottom = visual_bottom.max(cursor + block_height);
                cursor += block_height;
            }
            _ => {}
        }
    }
    (visual_top, visual_bottom)
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

pub fn extend_body_margins(
    page_size: &Size,
    margins: &PageMargins,
    header_height: f64,
    footer_height: f64,
) -> PageMargins {
    let header_distance = margins.header.unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let footer_distance = margins.footer.unwrap_or(DEFAULT_HF_DISTANCE_PX);
    let mut output = margins.clone();
    if header_height > margins.top - header_distance {
        output.top = margins.top.max(header_distance + header_height);
    }
    if footer_height > margins.bottom - footer_distance {
        output.bottom = margins.bottom.max(footer_distance + footer_height);
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
    use serde_json::json;

    use super::*;

    #[test]
    fn normalization_strips_inherited_spacing_and_suppresses_table_tail() {
        let blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
            {"kind": "table", "id": "t", "rows": []},
            {"kind": "paragraph", "id": "p", "runs": [], "attrs": {"spacing": {"before": 8, "after": 9}}}
        ]))
        .unwrap();

        let normalized = normalize_header_footer_blocks(blocks);
        let LayoutBlock::Paragraph(paragraph) = &normalized[1] else {
            panic!("paragraph expected");
        };
        let attrs = paragraph.attrs.as_ref().unwrap();
        assert_eq!(attrs.spacing.as_ref().unwrap().before, None);
        assert_eq!(attrs.spacing.as_ref().unwrap().after, None);
        assert_eq!(attrs.suppress_empty_paragraph_height, Some(true));
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
}
