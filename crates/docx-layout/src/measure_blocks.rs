use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::table_grid::{resolve_cell_grid, resolve_table_column_widths, resolve_table_width_px};
use crate::types::{
    BlockExtent, ChartExtent, ImageExtent, LayoutBlock, ParagraphBlock, ParagraphExtent, Run,
    ShapeBlock, ShapeExtent, TableBlock, TableCellExtent, TableExtent, TableRowExtent,
    TextBoxExtent,
};

const DEFAULT_CELL_PADDING_X: f64 = 7.0;
const DEFAULT_CELL_PADDING_Y: f64 = 0.0;

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
            let left = margins.map_or(9.6, |value| value.left);
            let right = margins.map_or(9.6, |value| value.right);
            let top = margins.map_or(4.8, |value| value.top);
            let bottom = margins.map_or(4.8, |value| value.bottom);
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

fn rotation_bound(bounds: &Option<Value>, field: &str) -> Option<f64> {
    bounds.as_ref()?.get(field)?.as_f64()
}

pub(crate) fn measure_paragraph(
    paragraph: &ParagraphBlock,
    content_width: f64,
    config: &MeasurementConfig,
) -> Result<ParagraphExtent, String> {
    if !content_width.is_finite() || content_width <= 0.0 {
        return Err("invalid: paragraph content width must be positive and finite".to_owned());
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
    let extent = crate::measure_paragraph_json_resident(&envelope.to_string())?;
    serde_json::from_str(&extent).map_err(|error| format!("parse paragraph extent: {error}"))
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
}
