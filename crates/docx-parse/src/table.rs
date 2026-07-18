//! Document-story table models and incumbent-compatible property leaves.
//!
//! The similarly named helpers in `formatting` intentionally parse table
//! properties inside style definitions. The live TypeScript parser has a
//! separate document-table path with different malformed/default handling;
//! S7 mirrors that path here.

use serde::{Deserialize, Serialize};

use crate::block::BlockContent;
use crate::borders::{Borders, parse_border_spec};
use crate::formatting::{
    CellMargins, ConditionalFormatStyle, FloatingTableProperties, TableCellFormatting,
    TableFormatting, TableLook, TableMeasurement, TableRowFormatting,
};
use crate::inline::{InlineNode, RunContent};
use crate::paragraph::{ParagraphContent, TrackedChangeInfo};
use crate::scalars::{ColorValue, ShadingProperties};
use crate::xml::{XmlElement, parse_javascript_integer_prefix};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Table {
    #[serde(rename = "type")]
    pub node_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<TableFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub property_changes: Option<Vec<TablePropertyChange>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_widths: Option<Vec<f64>>,
    pub rows: Vec<TableRow>,
}

impl Table {
    pub fn empty() -> Self {
        Self {
            node_type: "table".to_owned(),
            formatting: None,
            property_changes: None,
            column_widths: None,
            rows: Vec::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    #[serde(rename = "type")]
    pub node_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<TableRowFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub property_changes: Option<Vec<TableRowPropertyChange>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub structural_change: Option<TableStructuralChangeInfo>,
    pub cells: Vec<TableCell>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    #[serde(rename = "type")]
    pub node_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<TableCellFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub property_changes: Option<Vec<TableCellPropertyChange>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub structural_change: Option<TableStructuralChangeInfo>,
    pub content: Vec<BlockContent>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TablePropertyChange {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub previous_formatting: Option<TableFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_formatting: Option<TableFormatting>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRowPropertyChange {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub previous_formatting: Option<TableRowFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_formatting: Option<TableRowFormatting>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellPropertyChange {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub previous_formatting: Option<TableCellFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_formatting: Option<TableCellFormatting>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStructuralChangeInfo {
    #[serde(rename = "type")]
    pub node_type: String,
    pub info: TrackedChangeInfo,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_merge: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_merge_orig: Option<String>,
}

pub fn parse_table_measurement(element: Option<&XmlElement>) -> Option<TableMeasurement> {
    let element = element?;
    let value = element
        .parse_numeric_attribute(Some("w"), "w", 1.0)
        .unwrap_or(0.0);
    let kind = match element.attribute(Some("w"), "type") {
        Some(value @ ("auto" | "dxa" | "nil" | "pct")) => value,
        _ => "dxa",
    };
    Some(TableMeasurement {
        value,
        kind: kind.to_owned(),
    })
}

pub fn parse_document_table_borders(element: Option<&XmlElement>) -> Option<Borders> {
    let element = element?;
    let borders = Borders {
        top: parse_border_spec(element.child("w", "top")),
        bottom: parse_border_spec(element.child("w", "bottom")),
        left: parse_border_spec(element.child("w", "left")),
        right: parse_border_spec(element.child("w", "right")),
        inside_h: parse_border_spec(element.child("w", "insideH")),
        inside_v: parse_border_spec(element.child("w", "insideV")),
        start: parse_border_spec(element.child("w", "start")),
        end: parse_border_spec(element.child("w", "end")),
        tl2br: parse_border_spec(element.child("w", "tl2br")),
        tr2bl: parse_border_spec(element.child("w", "tr2bl")),
        ..Borders::default()
    };
    (borders != Borders::default()).then_some(borders)
}

pub fn parse_cell_margins(element: Option<&XmlElement>) -> Option<CellMargins> {
    let element = element?;
    let margins = CellMargins {
        top: parse_table_measurement(element.child("w", "top")),
        bottom: parse_table_measurement(element.child("w", "bottom")),
        left: parse_table_measurement(element.child("w", "left")),
        right: parse_table_measurement(element.child("w", "right")),
        start: parse_table_measurement(element.child("w", "start")),
        end: parse_table_measurement(element.child("w", "end")),
    };
    (margins != CellMargins::default()).then_some(margins)
}

pub fn parse_document_shading(element: Option<&XmlElement>) -> Option<ShadingProperties> {
    let element = element?;
    let mut shading = ShadingProperties::default();
    if let Some(fill) = element
        .attribute(Some("w"), "fill")
        .filter(|value| !value.is_empty() && *value != "auto")
    {
        shading.fill = Some(ColorValue {
            rgb: Some(fill.to_owned()),
            ..ColorValue::default()
        });
    }
    if let Some(theme_fill) = element
        .attribute(Some("w"), "themeFill")
        .filter(|value| !value.is_empty())
    {
        shading.fill = Some(ColorValue {
            theme_color: Some(theme_fill.to_owned()),
            theme_tint: element
                .attribute(Some("w"), "themeFillTint")
                .filter(|value| !value.is_empty())
                .map(str::to_owned),
            theme_shade: element
                .attribute(Some("w"), "themeFillShade")
                .filter(|value| !value.is_empty())
                .map(str::to_owned),
            ..ColorValue::default()
        });
    }
    if let Some(color) = element
        .attribute(Some("w"), "color")
        .filter(|value| !value.is_empty() && *value != "auto")
    {
        shading.color = Some(ColorValue {
            rgb: Some(color.to_owned()),
            ..ColorValue::default()
        });
    }
    if let Some(theme_color) = element
        .attribute(Some("w"), "themeColor")
        .filter(|value| !value.is_empty())
    {
        let color = shading.color.get_or_insert_with(ColorValue::default);
        color.theme_color = Some(theme_color.to_owned());
        color.theme_tint = element.attribute(Some("w"), "themeTint").map(str::to_owned);
        color.theme_shade = element
            .attribute(Some("w"), "themeShade")
            .map(str::to_owned);
    }
    shading.pattern = element
        .attribute(Some("w"), "val")
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    (shading != ShadingProperties::default()).then_some(shading)
}

pub fn parse_document_table_look(element: Option<&XmlElement>) -> Option<TableLook> {
    let element = element?;
    let mut look = TableLook::default();
    if let Some(raw) = element
        .attribute(Some("w"), "val")
        .filter(|value| !value.is_empty())
    {
        look.value = Some(raw.chars().take(8).collect());
        if let Some(bits) = parse_javascript_hex_prefix(raw) {
            look.first_row = Some(bits & 0x0020 != 0);
            look.last_row = Some(bits & 0x0040 != 0);
            look.first_column = Some(bits & 0x0080 != 0);
            look.last_column = Some(bits & 0x0100 != 0);
            look.no_h_band = Some(bits & 0x0200 != 0);
            look.no_v_band = Some(bits & 0x0400 != 0);
        }
    }
    for (attribute, slot) in [
        ("firstRow", &mut look.first_row),
        ("lastRow", &mut look.last_row),
        ("firstColumn", &mut look.first_column),
        ("lastColumn", &mut look.last_column),
        ("noHBand", &mut look.no_h_band),
        ("noVBand", &mut look.no_v_band),
    ] {
        if let Some(raw) = element.attribute(Some("w"), attribute) {
            *slot = Some(!matches_ci(raw, &["0", "false", "off"]));
        }
    }
    (look != TableLook::default()).then_some(look)
}

pub fn parse_conditional_format_style(
    element: Option<&XmlElement>,
) -> Option<ConditionalFormatStyle> {
    let element = element?;
    let mut style = ConditionalFormatStyle::default();
    if let Some(raw) = element
        .attribute(Some("w"), "val")
        .filter(|value| value.chars().count() == 12)
    {
        let bits: Vec<_> = raw.chars().collect();
        style.first_row = Some(bits[0] == '1');
        style.last_row = Some(bits[1] == '1');
        style.first_column = Some(bits[2] == '1');
        style.last_column = Some(bits[3] == '1');
        style.odd_v_band = Some(bits[4] == '1');
        style.even_v_band = Some(bits[5] == '1');
        style.odd_h_band = Some(bits[6] == '1');
        style.even_h_band = Some(bits[7] == '1');
        style.nw_cell = Some(bits[8] == '1');
        style.ne_cell = Some(bits[9] == '1');
        style.sw_cell = Some(bits[10] == '1');
        style.se_cell = Some(bits[11] == '1');
    }
    for (attribute, slot) in [
        ("firstRow", &mut style.first_row),
        ("lastRow", &mut style.last_row),
        ("firstColumn", &mut style.first_column),
        ("lastColumn", &mut style.last_column),
        ("oddHBand", &mut style.odd_h_band),
        ("evenHBand", &mut style.even_h_band),
        ("oddVBand", &mut style.odd_v_band),
        ("evenVBand", &mut style.even_v_band),
        ("firstRowFirstColumn", &mut style.nw_cell),
        ("firstRowLastColumn", &mut style.ne_cell),
        ("lastRowFirstColumn", &mut style.sw_cell),
        ("lastRowLastColumn", &mut style.se_cell),
    ] {
        if let Some(raw) = element.attribute(Some("w"), attribute) {
            *slot = Some(!matches_ci(raw, &["0", "false", "off"]));
        }
    }
    (style != ConditionalFormatStyle::default()).then_some(style)
}

pub fn parse_floating_table_properties(
    element: Option<&XmlElement>,
) -> Option<FloatingTableProperties> {
    let element = element?;
    let floating = FloatingTableProperties {
        horz_anchor: attribute_in(element, "horzAnchor", &["margin", "page", "text"]),
        vert_anchor: attribute_in(element, "vertAnchor", &["margin", "page", "text"]),
        tblp_x: element.parse_numeric_attribute(Some("w"), "tblpX", 1.0),
        tblp_x_spec: attribute_nonempty(element, "tblpXSpec"),
        tblp_y: element.parse_numeric_attribute(Some("w"), "tblpY", 1.0),
        tblp_y_spec: attribute_nonempty(element, "tblpYSpec"),
        top_from_text: element.parse_numeric_attribute(Some("w"), "topFromText", 1.0),
        bottom_from_text: element.parse_numeric_attribute(Some("w"), "bottomFromText", 1.0),
        left_from_text: element.parse_numeric_attribute(Some("w"), "leftFromText", 1.0),
        right_from_text: element.parse_numeric_attribute(Some("w"), "rightFromText", 1.0),
    };
    (floating != FloatingTableProperties::default()).then_some(floating)
}

pub fn parse_document_table_properties(element: Option<&XmlElement>) -> Option<TableFormatting> {
    let element = element?;
    let mut formatting = TableFormatting::default();
    formatting.width = parse_table_measurement(element.child("w", "tblW"));
    formatting.justification = element
        .child("w", "jc")
        .and_then(|child| child.attribute(Some("w"), "val"))
        .and_then(|value| match value {
            "left" | "center" | "right" => Some(value.to_owned()),
            "start" => Some("left".to_owned()),
            _ => None,
        });
    formatting.cell_spacing = parse_table_measurement(element.child("w", "tblCellSpacing"));
    formatting.indent = parse_table_measurement(element.child("w", "tblInd"));
    formatting.borders = parse_document_table_borders(element.child("w", "tblBorders"));
    formatting.cell_margins = parse_cell_margins(element.child("w", "tblCellMar"));
    formatting.layout = child_attribute_in(element, "tblLayout", "type", &["fixed", "autofit"]);
    formatting.style_id = child_attribute_nonempty(element, "tblStyle", "val");
    formatting.style_row_band_size = child_number(element, "tblStyleRowBandSize", "val")
        .filter(|value| *value > 0.0 && *value <= 1024.0);
    formatting.style_col_band_size = child_number(element, "tblStyleColBandSize", "val")
        .filter(|value| *value > 0.0 && *value <= 1024.0);
    formatting.look = parse_document_table_look(element.child("w", "tblLook"));
    formatting.shading = parse_document_shading(element.child("w", "shd"));
    formatting.overlap = child_attribute_in(element, "tblOverlap", "val", &["never", "overlap"]);
    formatting.floating = parse_floating_table_properties(element.child("w", "tblpPr"));
    formatting.bidi = element
        .child("w", "bidiVisual")
        .filter(|child| child.parse_boolean("w"))
        .map(|_| true);
    (formatting != TableFormatting::default()).then_some(formatting)
}

pub fn parse_document_table_row_properties(
    element: Option<&XmlElement>,
) -> Option<TableRowFormatting> {
    let element = element?;
    let mut formatting = TableRowFormatting::default();
    if let Some(height) = element.child("w", "trHeight") {
        if let Some(value) = height
            .parse_numeric_attribute(Some("w"), "val", 1.0)
            .filter(|value| *value > 0.0)
        {
            formatting.height = Some(TableMeasurement {
                value,
                kind: "dxa".to_owned(),
            });
        }
        formatting.height_rule = height
            .attribute(Some("w"), "hRule")
            .filter(|value| matches!(*value, "auto" | "atLeast" | "exact"))
            .map(str::to_owned);
    }
    formatting.header = true_child(element, "tblHeader");
    formatting.cant_split = true_child(element, "cantSplit");
    formatting.justification =
        child_attribute_in(element, "jc", "val", &["left", "center", "right"]);
    formatting.hidden = true_child(element, "hidden");
    formatting.conditional_format = parse_conditional_format_style(element.child("w", "cnfStyle"));
    formatting.grid_before = child_number(element, "gridBefore", "val")
        .filter(|value| *value >= 0.0 && *value <= 16_384.0);
    formatting.grid_after = child_number(element, "gridAfter", "val")
        .filter(|value| *value >= 0.0 && *value <= 16_384.0);
    formatting.width_before = parse_table_measurement(element.child("w", "wBefore"));
    formatting.width_after = parse_table_measurement(element.child("w", "wAfter"));
    (formatting != TableRowFormatting::default()).then_some(formatting)
}

pub fn parse_document_table_cell_properties(
    element: Option<&XmlElement>,
) -> Option<TableCellFormatting> {
    let element = element?;
    let mut formatting = TableCellFormatting::default();
    formatting.width = parse_table_measurement(element.child("w", "tcW"));
    formatting.borders = parse_document_table_borders(element.child("w", "tcBorders"));
    formatting.margins = parse_cell_margins(element.child("w", "tcMar"));
    formatting.shading = parse_document_shading(element.child("w", "shd"));
    formatting.vertical_align =
        child_attribute_in(element, "vAlign", "val", &["top", "center", "bottom"]);
    formatting.text_direction = child_attribute_nonempty(element, "textDirection", "val");
    formatting.grid_span = child_number(element, "gridSpan", "val").filter(|value| *value > 1.0);
    formatting.v_merge = element.child("w", "vMerge").map(|merge| {
        if merge.attribute(Some("w"), "val") == Some("restart") {
            "restart"
        } else {
            "continue"
        }
        .to_owned()
    });
    formatting.fit_text = true_child(element, "tcFitText");
    formatting.no_wrap = true_child(element, "noWrap");
    formatting.hide_mark = true_child(element, "hideMark");
    formatting.conditional_format = parse_conditional_format_style(element.child("w", "cnfStyle"));
    (formatting != TableCellFormatting::default()).then_some(formatting)
}

pub fn parse_table_property_changes(
    element: Option<&XmlElement>,
    current: Option<&TableFormatting>,
) -> Option<Vec<TablePropertyChange>> {
    let changes: Vec<_> = element?
        .children_named("w", "tblPrChange")
        .filter_map(|change| {
            let previous = parse_document_table_properties(change.child("w", "tblPr"));
            let current = current.cloned();
            (previous.is_some() || current.is_some()).then(|| TablePropertyChange {
                node_type: "tablePropertyChange".to_owned(),
                info: parse_change_info(change),
                previous_formatting: previous,
                current_formatting: current,
            })
        })
        .collect();
    (!changes.is_empty()).then_some(changes)
}

pub fn parse_table_row_property_changes(
    element: Option<&XmlElement>,
    current: Option<&TableRowFormatting>,
) -> Option<Vec<TableRowPropertyChange>> {
    let changes: Vec<_> = element?
        .children_named("w", "trPrChange")
        .filter_map(|change| {
            let previous = parse_document_table_row_properties(change.child("w", "trPr"));
            let current = current.cloned();
            (previous.is_some() || current.is_some()).then(|| TableRowPropertyChange {
                node_type: "tableRowPropertyChange".to_owned(),
                info: parse_change_info(change),
                previous_formatting: previous,
                current_formatting: current,
            })
        })
        .collect();
    (!changes.is_empty()).then_some(changes)
}

pub fn parse_table_cell_property_changes(
    element: Option<&XmlElement>,
    current: Option<&TableCellFormatting>,
) -> Option<Vec<TableCellPropertyChange>> {
    let changes: Vec<_> = element?
        .children_named("w", "tcPrChange")
        .filter_map(|change| {
            let previous = parse_document_table_cell_properties(change.child("w", "tcPr"));
            let current = current.cloned();
            (previous.is_some() || current.is_some()).then(|| TableCellPropertyChange {
                node_type: "tableCellPropertyChange".to_owned(),
                info: parse_change_info(change),
                previous_formatting: previous,
                current_formatting: current,
            })
        })
        .collect();
    (!changes.is_empty()).then_some(changes)
}

pub fn parse_table_row_structural_change(
    element: Option<&XmlElement>,
) -> Option<TableStructuralChangeInfo> {
    let element = element?;
    if let Some(change) = element.child("w", "ins") {
        return Some(structural_change("tableRowInsertion", change));
    }
    element
        .child("w", "del")
        .map(|change| structural_change("tableRowDeletion", change))
}

pub fn parse_table_cell_structural_change(
    element: Option<&XmlElement>,
) -> Option<TableStructuralChangeInfo> {
    let element = element?;
    if let Some(change) = element.child("w", "cellIns") {
        return Some(structural_change("tableCellInsertion", change));
    }
    if let Some(change) = element.child("w", "cellDel") {
        return Some(structural_change("tableCellDeletion", change));
    }
    element.child("w", "cellMerge").map(|change| {
        let mut parsed = structural_change("tableCellMerge", change);
        parsed.v_merge = annotation_v_merge(change.attribute(Some("w"), "vMerge"));
        parsed.v_merge_orig = annotation_v_merge(change.attribute(Some("w"), "vMergeOrig"));
        parsed
    })
}

pub fn parse_table_grid(element: Option<&XmlElement>) -> Option<Vec<f64>> {
    let widths: Vec<_> = element?
        .children_named("w", "gridCol")
        .map(|column| {
            column
                .parse_numeric_attribute(Some("w"), "w", 1.0)
                .unwrap_or(0.0)
        })
        .collect();
    (!widths.is_empty() && widths.iter().any(|width| *width > 0.0)).then_some(widths)
}

pub fn infer_implicit_single_cell_row_spans(table: &mut Table) {
    let max_columns = table.rows.iter().map(row_grid_span).fold(
        table
            .column_widths
            .as_ref()
            .map_or(0.0, |grid| grid.len() as f64),
        f64::max,
    );
    if max_columns <= 1.0 {
        return;
    }
    for row in &mut table.rows {
        if row.cells.len() != 1 {
            continue;
        }
        let cell = &mut row.cells[0];
        let formatting = cell
            .formatting
            .get_or_insert_with(TableCellFormatting::default);
        let current_span = formatting.grid_span.unwrap_or(1.0);
        if current_span >= max_columns
            || formatting.v_merge.is_some()
            || formatting.grid_span.is_some()
        {
            continue;
        }
        formatting.grid_span = Some(max_columns);
    }
}

pub fn get_table_column_count(table: &Table) -> usize {
    if let Some(grid) = &table.column_widths
        && !grid.is_empty()
    {
        return grid.len();
    }
    table.rows.first().map_or(0, |row| {
        row_grid_span(row).max(0.0).min(usize::MAX as f64) as usize
    })
}

pub fn get_table_row_count(table: &Table) -> usize {
    table.rows.len()
}

pub fn is_cell_merge_continuation(cell: &TableCell) -> bool {
    cell.formatting
        .as_ref()
        .and_then(|formatting| formatting.v_merge.as_deref())
        == Some("continue")
}

pub fn is_cell_merge_start(cell: &TableCell) -> bool {
    cell.formatting
        .as_ref()
        .and_then(|formatting| formatting.v_merge.as_deref())
        == Some("restart")
}

pub fn is_cell_horizontally_merged(cell: &TableCell) -> bool {
    cell.formatting
        .as_ref()
        .and_then(|formatting| formatting.grid_span)
        .unwrap_or(1.0)
        > 1.0
}

pub fn get_table_text(table: &Table) -> String {
    table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| {
                    cell.content
                        .iter()
                        .filter_map(|block| match block {
                            BlockContent::Paragraph(paragraph) => Some(
                                paragraph
                                    .content
                                    .iter()
                                    .filter_map(|content| match content {
                                        ParagraphContent::Inline(InlineNode::Run(run)) => Some(
                                            run.content
                                                .iter()
                                                .filter_map(|content| match content {
                                                    RunContent::Text { text, .. } => {
                                                        Some(text.as_str())
                                                    }
                                                    _ => None,
                                                })
                                                .collect::<String>(),
                                        ),
                                        _ => None,
                                    })
                                    .collect::<String>(),
                            ),
                            _ => None,
                        })
                        .collect::<Vec<_>>()
                        .join("\n")
                })
                .collect::<Vec<_>>()
                .join("\t")
        })
        .collect::<Vec<_>>()
        .join("\n")
}

pub fn has_header_row(table: &Table) -> bool {
    table.rows.first().is_some_and(|row| {
        row.formatting
            .as_ref()
            .and_then(|formatting| formatting.header)
            == Some(true)
    })
}

pub fn get_header_rows(table: &Table) -> Vec<&TableRow> {
    table
        .rows
        .iter()
        .filter(|row| {
            row.formatting
                .as_ref()
                .and_then(|formatting| formatting.header)
                == Some(true)
        })
        .collect()
}

pub fn is_floating_table(table: &Table) -> bool {
    table
        .formatting
        .as_ref()
        .and_then(|formatting| formatting.floating.as_ref())
        .is_some()
}

fn row_grid_span(row: &TableRow) -> f64 {
    row.cells
        .iter()
        .map(|cell| {
            cell.formatting
                .as_ref()
                .and_then(|formatting| formatting.grid_span)
                .unwrap_or(1.0)
        })
        .sum()
}

fn parse_change_info(element: &XmlElement) -> TrackedChangeInfo {
    let id = element
        .attribute(Some("w"), "id")
        .and_then(parse_javascript_integer_prefix)
        .filter(|id| id.fract() == 0.0 && *id >= 0.0)
        .unwrap_or(0.0);
    let author = element
        .attribute(Some("w"), "author")
        .unwrap_or_default()
        .trim();
    let date = element
        .attribute(Some("w"), "date")
        .unwrap_or_default()
        .trim();
    TrackedChangeInfo {
        id,
        author: if author.is_empty() { "Unknown" } else { author }.to_owned(),
        date: (!date.is_empty()).then(|| date.to_owned()),
    }
}

fn structural_change(node_type: &str, element: &XmlElement) -> TableStructuralChangeInfo {
    TableStructuralChangeInfo {
        node_type: node_type.to_owned(),
        info: parse_change_info(element),
        v_merge: None,
        v_merge_orig: None,
    }
}

fn annotation_v_merge(value: Option<&str>) -> Option<String> {
    value
        .filter(|value| matches!(*value, "rest" | "cont"))
        .map(str::to_owned)
}

fn true_child(parent: &XmlElement, name: &str) -> Option<bool> {
    parent
        .child("w", name)
        .filter(|element| element.parse_boolean("w"))
        .map(|_| true)
}

fn child_number(parent: &XmlElement, child: &str, attribute: &str) -> Option<f64> {
    parent
        .child("w", child)
        .and_then(|element| element.parse_numeric_attribute(Some("w"), attribute, 1.0))
}

fn child_attribute_nonempty(parent: &XmlElement, child: &str, attribute: &str) -> Option<String> {
    parent
        .child("w", child)?
        .attribute(Some("w"), attribute)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

fn child_attribute_in(
    parent: &XmlElement,
    child: &str,
    attribute: &str,
    allowed: &[&str],
) -> Option<String> {
    let value = parent.child("w", child)?.attribute(Some("w"), attribute)?;
    allowed.contains(&value).then(|| value.to_owned())
}

fn attribute_nonempty(element: &XmlElement, attribute: &str) -> Option<String> {
    element
        .attribute(Some("w"), attribute)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

fn attribute_in(element: &XmlElement, attribute: &str, allowed: &[&str]) -> Option<String> {
    let value = element.attribute(Some("w"), attribute)?;
    allowed.contains(&value).then(|| value.to_owned())
}

fn matches_ci(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

fn parse_javascript_hex_prefix(raw: &str) -> Option<u32> {
    let trimmed = raw.trim_start();
    let unsigned = trimmed
        .strip_prefix('+')
        .or_else(|| trimmed.strip_prefix('-'))
        .unwrap_or(trimmed);
    let unsigned = unsigned
        .strip_prefix("0x")
        .or_else(|| unsigned.strip_prefix("0X"))
        .unwrap_or(unsigned);
    let digits: String = unsigned
        .chars()
        .take_while(char::is_ascii_hexdigit)
        .collect();
    if digits.is_empty() {
        None
    } else {
        u64::from_str_radix(&digits.chars().take(16).collect::<String>(), 16)
            .ok()
            .map(|value| value as u32)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(xml.as_bytes(), "table.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn pins_document_measurement_grid_span_and_vmerge_normalization() {
        let properties = root(
            r#"<w:tcPr xmlns:w="w"><w:tcW w:type="bogus"/><w:gridSpan w:val="1"/><w:vMerge w:val="bogus"/><w:noWrap w:val="0"/></w:tcPr>"#,
        );
        let parsed = parse_document_table_cell_properties(Some(&properties)).unwrap();
        assert_eq!(
            parsed.width,
            Some(TableMeasurement {
                value: 0.0,
                kind: "dxa".to_owned(),
            })
        );
        assert_eq!(parsed.grid_span, None);
        assert_eq!(parsed.v_merge.as_deref(), Some("continue"));
        assert_eq!(parsed.no_wrap, None);
    }

    #[test]
    fn pins_floating_look_conditional_borders_and_shading() {
        let properties = root(
            r#"<w:tblPr xmlns:w="w">
              <w:jc w:val="start"/><w:tblOverlap w:val="never"/>
              <w:tblpPr w:horzAnchor="page" w:vertAnchor="text" w:tblpX="-20px" w:tblpXSpec="outside" w:topFromText="40"/>
              <w:tblLook w:val="04A0" w:firstRow="0"/>
              <w:tblBorders><w:left w:val="single"/><w:start w:val="double"/><w:tl2br w:val="dashed"/></w:tblBorders>
              <w:shd w:fill="FF0000" w:themeFill="accent1" w:themeFillTint="80"/>
            </w:tblPr>"#,
        );
        let parsed = parse_document_table_properties(Some(&properties)).unwrap();
        assert_eq!(parsed.justification.as_deref(), Some("left"));
        assert_eq!(parsed.overlap.as_deref(), Some("never"));
        let floating = parsed.floating.unwrap();
        assert_eq!(floating.tblp_x, Some(-20.0));
        assert_eq!(floating.tblp_x_spec.as_deref(), Some("outside"));
        let borders = parsed.borders.unwrap();
        assert_eq!(borders.left.unwrap().style, "single");
        assert_eq!(borders.start.unwrap().style, "double");
        assert_eq!(borders.tl2br.unwrap().style, "dashed");
        let look = parsed.look.unwrap();
        assert_eq!(look.first_row, Some(false));
        assert_eq!(look.no_v_band, Some(true));
        let fill = parsed.shading.unwrap().fill.unwrap();
        assert_eq!(fill.rgb, None);
        assert_eq!(fill.theme_color.as_deref(), Some("accent1"));

        let conditional = root(
            r#"<w:cnfStyle xmlns:w="w" w:val="100000001000" w:firstRow="off" w:lastRowLastColumn="1"/>"#,
        );
        let conditional = parse_conditional_format_style(Some(&conditional)).unwrap();
        assert_eq!(conditional.first_row, Some(false));
        assert_eq!(conditional.nw_cell, Some(true));
        assert_eq!(conditional.se_cell, Some(true));
    }

    #[test]
    fn pins_property_and_structural_change_precedence() {
        let row = root(
            r#"<w:trPr xmlns:w="w"><w:tblHeader/><w:ins w:id="7x" w:author=" A "/><w:del w:id="8"/><w:trPrChange w:id="bad"><w:trPr><w:hidden/></w:trPr></w:trPrChange></w:trPr>"#,
        );
        let current = parse_document_table_row_properties(Some(&row));
        let changes = parse_table_row_property_changes(Some(&row), current.as_ref()).unwrap();
        assert_eq!(changes[0].info.id, 0.0);
        assert_eq!(changes[0].info.author, "Unknown");
        assert_eq!(
            changes[0].previous_formatting.as_ref().unwrap().hidden,
            Some(true)
        );
        let structural = parse_table_row_structural_change(Some(&row)).unwrap();
        assert_eq!(structural.node_type, "tableRowInsertion");
        assert_eq!(structural.info.id, 7.0);
        assert_eq!(structural.info.author, "A");

        let cell = root(
            r#"<w:tcPr xmlns:w="w"><w:cellMerge w:id="3" w:vMerge="rest" w:vMergeOrig="invalid"/></w:tcPr>"#,
        );
        let structural = parse_table_cell_structural_change(Some(&cell)).unwrap();
        assert_eq!(structural.v_merge.as_deref(), Some("rest"));
        assert_eq!(structural.v_merge_orig, None);
    }

    #[test]
    fn preserves_zero_columns_but_omits_an_empty_or_all_nonpositive_grid() {
        let mixed = root(
            r#"<w:tblGrid xmlns:w="w"><w:gridCol/><w:gridCol w:w="1200px"/><w:gridCol w:w="-1"/></w:tblGrid>"#,
        );
        assert_eq!(
            parse_table_grid(Some(&mixed)),
            Some(vec![0.0, 1200.0, -1.0])
        );
        let empty = root(r#"<w:tblGrid xmlns:w="w"><w:gridCol/><w:gridCol w:w="0"/></w:tblGrid>"#);
        assert_eq!(parse_table_grid(Some(&empty)), None);
    }
}
