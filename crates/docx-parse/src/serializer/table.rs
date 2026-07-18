//! Table, row, cell, and table-property serializers.

use crate::block::BlockContent;
use crate::borders::Borders;
use crate::formatting::{
    CellMargins, FloatingTableProperties, TableCellFormatting, TableFormatting, TableLook,
    TableMeasurement, TableRowFormatting,
};
use crate::paragraph::TrackedChangeInfo;
use crate::table::{
    Table, TableCell, TableCellPropertyChange, TablePropertyChange, TableRow,
    TableRowPropertyChange, TableStructuralChangeInfo,
};
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::foundation::{
    BorderSide, serialize_conditional_format_style, serialize_table_grid, write_border,
};
use super::paragraph::serialize_paragraph;
use super::run::{
    append_generated, nonempty, nonempty_trimmed, normalized_tracked_id, write_shading,
};
use super::xml_writer::{XmlWriter, int_attr};

pub fn serialize_table(
    table: &Table,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let properties =
        serialize_table_formatting(table.formatting.as_ref(), table.property_changes.as_deref());
    let grid = serialize_table_grid(table);
    let mut writer = XmlWriter::with_capacity(1024);
    writer.start_element("w:tbl");
    if !properties.is_empty() {
        append_generated(&mut writer, &properties);
    } else if !grid.is_empty() {
        writer.start_element("w:tblPr").end_element();
    }
    append_generated(&mut writer, &grid);
    for row in &table.rows {
        append_generated(&mut writer, &serialize_table_row(row, context)?);
    }
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_table_formatting(
    formatting: Option<&TableFormatting>,
    changes: Option<&[TablePropertyChange]>,
) -> String {
    let mut body = XmlWriter::with_capacity(512);
    if let Some(formatting) = formatting {
        if let Some(value) = nonempty(formatting.style_id.as_deref()) {
            empty_attr(&mut body, "w:tblStyle", "w:val", value);
        }
        write_floating(&mut body, formatting.floating.as_ref());
        if let Some(value) = nonempty(formatting.overlap.as_deref()) {
            empty_attr(&mut body, "w:tblOverlap", "w:val", value);
        }
        if formatting.bidi == Some(true) {
            body.start_element("w:bidiVisual").end_element();
        }
        write_measurement(&mut body, formatting.width.as_ref(), "w:tblW");
        if let Some(value) = nonempty(formatting.justification.as_deref()) {
            empty_attr(&mut body, "w:jc", "w:val", value);
        }
        write_measurement(
            &mut body,
            formatting.cell_spacing.as_ref(),
            "w:tblCellSpacing",
        );
        write_measurement(&mut body, formatting.indent.as_ref(), "w:tblInd");
        write_borders(&mut body, formatting.borders.as_ref(), "w:tblBorders");
        write_shading(&mut body, formatting.shading.as_ref());
        if let Some(value) = nonempty(formatting.layout.as_deref()) {
            empty_attr(&mut body, "w:tblLayout", "w:type", value);
        }
        write_margins(&mut body, formatting.cell_margins.as_ref(), "w:tblCellMar");
        write_table_look(&mut body, formatting.look.as_ref());
    }
    if let Some(change) = changes.and_then(|changes| changes.first()) {
        append_generated(&mut body, &serialize_table_property_change(change));
    }
    wrap_properties("w:tblPr", body.finish())
}

pub fn serialize_table_row(
    row: &TableRow,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let properties = serialize_table_row_formatting(
        row.formatting.as_ref(),
        row.property_changes.as_deref(),
        row.structural_change.as_ref(),
    );
    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:tr");
    append_generated(&mut writer, &properties);
    for cell in &row.cells {
        append_generated(&mut writer, &serialize_table_cell(cell, context)?);
    }
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_table_row_formatting(
    formatting: Option<&TableRowFormatting>,
    changes: Option<&[TableRowPropertyChange]>,
    structural: Option<&TableStructuralChangeInfo>,
) -> String {
    let mut body = XmlWriter::with_capacity(320);
    if let Some(formatting) = formatting {
        let conditional =
            serialize_conditional_format_style(formatting.conditional_format.as_ref());
        append_generated(&mut body, &conditional);
        if formatting.cant_split == Some(true) {
            body.start_element("w:cantSplit").end_element();
        }
        if formatting.header == Some(true) {
            body.start_element("w:tblHeader").end_element();
        }
        if let Some(height) = formatting.height.as_ref() {
            body.start_element("w:trHeight")
                .attribute("w:val", &int_attr(Some(height.value)));
            if let Some(rule) = nonempty(formatting.height_rule.as_deref()) {
                body.attribute("w:hRule", rule);
            }
            body.end_element();
        }
        if let Some(value) = nonempty(formatting.justification.as_deref()) {
            empty_attr(&mut body, "w:jc", "w:val", value);
        }
        if formatting.hidden == Some(true) {
            body.start_element("w:hidden").end_element();
        }
    }
    if let Some(structural) = structural {
        let element = match structural.node_type.as_str() {
            "tableRowInsertion" => Some("w:ins"),
            "tableRowDeletion" => Some("w:del"),
            _ => None,
        };
        if let Some(element) = element {
            write_tracked_empty(&mut body, element, &structural.info);
        }
    }
    if let Some(change) = changes.and_then(|changes| changes.first()) {
        append_generated(&mut body, &serialize_table_row_property_change(change));
    }
    wrap_properties("w:trPr", body.finish())
}

pub fn serialize_table_cell(
    cell: &TableCell,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let properties = serialize_table_cell_formatting(
        cell.formatting.as_ref(),
        cell.property_changes.as_deref(),
        cell.structural_change.as_ref(),
    );
    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:tc");
    append_generated(&mut writer, &properties);
    let mut emitted = false;
    for block in &cell.content {
        match block {
            BlockContent::Paragraph(paragraph) => {
                append_generated(&mut writer, &serialize_paragraph(paragraph, context)?);
                emitted = true;
            }
            BlockContent::Table(table) => {
                append_generated(&mut writer, &serialize_table(table, context)?);
                emitted = true;
            }
            // Pinned incumbent quirk: tableSerializer's cell dispatcher only
            // recognizes paragraph/table even though the shared story model
            // also permits a block SDT here.
            BlockContent::BlockSdt(_) => {}
        }
    }
    if !emitted {
        writer.start_element("w:p").end_element();
    }
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_table_cell_formatting(
    formatting: Option<&TableCellFormatting>,
    changes: Option<&[TableCellPropertyChange]>,
    structural: Option<&TableStructuralChangeInfo>,
) -> String {
    let mut body = XmlWriter::with_capacity(384);
    if let Some(formatting) = formatting {
        append_generated(
            &mut body,
            &serialize_conditional_format_style(formatting.conditional_format.as_ref()),
        );
        write_measurement(&mut body, formatting.width.as_ref(), "w:tcW");
        if let Some(span) = formatting.grid_span.filter(|span| *span > 1.0) {
            empty_attr(&mut body, "w:gridSpan", "w:val", &int_attr(Some(span)));
        }
        if let Some(merge) = nonempty(formatting.v_merge.as_deref()) {
            if merge == "restart" {
                empty_attr(&mut body, "w:vMerge", "w:val", "restart");
            } else {
                body.start_element("w:vMerge").end_element();
            }
        }
        write_borders(&mut body, formatting.borders.as_ref(), "w:tcBorders");
        write_shading(&mut body, formatting.shading.as_ref());
        if formatting.no_wrap == Some(true) {
            body.start_element("w:noWrap").end_element();
        }
        write_margins(&mut body, formatting.margins.as_ref(), "w:tcMar");
        if let Some(value) = nonempty(formatting.text_direction.as_deref()) {
            empty_attr(&mut body, "w:textDirection", "w:val", value);
        }
        if formatting.fit_text == Some(true) {
            body.start_element("w:tcFitText").end_element();
        }
        if let Some(value) = nonempty(formatting.vertical_align.as_deref()) {
            empty_attr(&mut body, "w:vAlign", "w:val", value);
        }
        if formatting.hide_mark == Some(true) {
            body.start_element("w:hideMark").end_element();
        }
    }
    if let Some(structural) = structural {
        match structural.node_type.as_str() {
            "tableCellInsertion" => write_tracked_empty(&mut body, "w:cellIns", &structural.info),
            "tableCellDeletion" => write_tracked_empty(&mut body, "w:cellDel", &structural.info),
            "tableCellMerge" => write_cell_merge(&mut body, structural),
            _ => {}
        }
    }
    if let Some(change) = changes.and_then(|changes| changes.first()) {
        append_generated(&mut body, &serialize_table_cell_property_change(change));
    }
    wrap_properties("w:tcPr", body.finish())
}

fn serialize_table_property_change(change: &TablePropertyChange) -> String {
    let previous = nonempty_properties(
        serialize_table_formatting(change.previous_formatting.as_ref(), None),
        "w:tblPr",
    );
    tracked_container("w:tblPrChange", &change.info, &previous)
}

fn serialize_table_row_property_change(change: &TableRowPropertyChange) -> String {
    let previous = nonempty_properties(
        serialize_table_row_formatting(change.previous_formatting.as_ref(), None, None),
        "w:trPr",
    );
    tracked_container("w:trPrChange", &change.info, &previous)
}

fn serialize_table_cell_property_change(change: &TableCellPropertyChange) -> String {
    let previous = nonempty_properties(
        serialize_table_cell_formatting(change.previous_formatting.as_ref(), None, None),
        "w:tcPr",
    );
    tracked_container("w:tcPrChange", &change.info, &previous)
}

fn tracked_container(element: &'static str, info: &TrackedChangeInfo, child: &str) -> String {
    let mut writer = XmlWriter::with_capacity(child.len() + 128);
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    append_generated(&mut writer, child);
    writer.end_element();
    writer.finish()
}

fn nonempty_properties(xml: String, element: &str) -> String {
    if xml.is_empty() {
        format!("<{element}/>")
    } else {
        xml
    }
}

fn write_tracked_empty(writer: &mut XmlWriter, element: &'static str, info: &TrackedChangeInfo) {
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    writer.end_element();
}

fn write_cell_merge(writer: &mut XmlWriter, structural: &TableStructuralChangeInfo) {
    let info = &structural.info;
    writer
        .start_element("w:cellMerge")
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    if let Some(value) = nonempty(structural.v_merge.as_deref()) {
        writer.attribute("w:vMerge", value);
    }
    if let Some(value) = nonempty(structural.v_merge_orig.as_deref()) {
        writer.attribute("w:vMergeOrig", value);
    }
    writer.end_element();
}

fn write_measurement(
    writer: &mut XmlWriter,
    measurement: Option<&TableMeasurement>,
    element: &'static str,
) {
    let Some(measurement) = measurement else {
        return;
    };
    writer
        .start_element(element)
        .attribute("w:w", &int_attr(Some(measurement.value)))
        .attribute("w:type", &measurement.kind)
        .end_element();
}

fn write_borders(writer: &mut XmlWriter, borders: Option<&Borders>, element: &'static str) {
    let Some(borders) = borders else {
        return;
    };
    let sides = [
        (borders.top.as_ref(), BorderSide::Top),
        (borders.left.as_ref(), BorderSide::Left),
        (borders.bottom.as_ref(), BorderSide::Bottom),
        (borders.right.as_ref(), BorderSide::Right),
        (borders.inside_h.as_ref(), BorderSide::InsideH),
        (borders.inside_v.as_ref(), BorderSide::InsideV),
    ];
    if !sides.iter().any(|(border, _)| border.is_some()) {
        return;
    }
    writer.start_element(element);
    for (border, side) in sides {
        if let Some(border) = border {
            write_border(writer, border, side);
        }
    }
    writer.end_element();
}

fn write_margins(writer: &mut XmlWriter, margins: Option<&CellMargins>, element: &'static str) {
    let Some(margins) = margins else {
        return;
    };
    let values = [
        (margins.top.as_ref(), "w:top"),
        (margins.left.as_ref(), "w:left"),
        (margins.bottom.as_ref(), "w:bottom"),
        (margins.right.as_ref(), "w:right"),
    ];
    if !values.iter().any(|(value, _)| value.is_some()) {
        return;
    }
    writer.start_element(element);
    for (value, name) in values {
        write_measurement(writer, value, name);
    }
    writer.end_element();
}

fn write_table_look(writer: &mut XmlWriter, look: Option<&TableLook>) {
    let Some(look) = look else {
        return;
    };
    let values = [
        (look.first_row, "w:firstRow"),
        (look.last_row, "w:lastRow"),
        (look.first_column, "w:firstColumn"),
        (look.last_column, "w:lastColumn"),
        (look.no_h_band, "w:noHBand"),
        (look.no_v_band, "w:noVBand"),
    ];
    if !values.iter().any(|(value, _)| *value == Some(true)) {
        return;
    }
    writer.start_element("w:tblLook");
    for (_, name) in values.into_iter().filter(|(value, _)| *value == Some(true)) {
        writer.attribute(name, "1");
    }
    writer.end_element();
}

fn write_floating(writer: &mut XmlWriter, floating: Option<&FloatingTableProperties>) {
    let Some(floating) = floating else {
        return;
    };
    let has_attributes = nonempty(floating.horz_anchor.as_deref()).is_some()
        || nonempty(floating.vert_anchor.as_deref()).is_some()
        || floating.tblp_x.is_some()
        || nonempty(floating.tblp_x_spec.as_deref()).is_some()
        || floating.tblp_y.is_some()
        || nonempty(floating.tblp_y_spec.as_deref()).is_some()
        || floating.top_from_text.is_some()
        || floating.bottom_from_text.is_some()
        || floating.left_from_text.is_some()
        || floating.right_from_text.is_some();
    if !has_attributes {
        return;
    }
    writer.start_element("w:tblpPr");
    optional_attr(writer, "w:horzAnchor", floating.horz_anchor.as_deref());
    optional_attr(writer, "w:vertAnchor", floating.vert_anchor.as_deref());
    optional_int(writer, "w:tblpX", floating.tblp_x);
    optional_attr(writer, "w:tblpXSpec", floating.tblp_x_spec.as_deref());
    optional_int(writer, "w:tblpY", floating.tblp_y);
    optional_attr(writer, "w:tblpYSpec", floating.tblp_y_spec.as_deref());
    optional_int(writer, "w:topFromText", floating.top_from_text);
    optional_int(writer, "w:bottomFromText", floating.bottom_from_text);
    optional_int(writer, "w:leftFromText", floating.left_from_text);
    optional_int(writer, "w:rightFromText", floating.right_from_text);
    writer.end_element();
}

fn optional_attr(writer: &mut XmlWriter, name: &'static str, value: Option<&str>) {
    if let Some(value) = nonempty(value) {
        writer.attribute(name, value);
    }
}

fn optional_int(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer.attribute(name, &int_attr(value));
    }
}

fn empty_attr(writer: &mut XmlWriter, element: &'static str, name: &'static str, value: &str) {
    writer
        .start_element(element)
        .attribute(name, value)
        .end_element();
}

fn wrap_properties(element: &str, body: String) -> String {
    if body.is_empty() {
        String::new()
    } else {
        format!("<{element}>{body}</{element}>")
    }
}

#[cfg(test)]
mod tests {
    use crate::formatting::{TableCellFormatting, TableMeasurement};
    use crate::serializer::s10::SerializerDeterminism;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    #[test]
    fn table_bytes_pin_required_grid_empty_cell_and_escape_properties() {
        let table = Table {
            node_type: "table".to_owned(),
            formatting: Some(TableFormatting {
                style_id: Some("bad\"/><evil&".to_owned()),
                width: Some(TableMeasurement {
                    value: 1008.0000000000001,
                    kind: "dxa".to_owned(),
                }),
                ..TableFormatting::default()
            }),
            property_changes: None,
            column_widths: Some(vec![1008.0000000000001]),
            rows: vec![TableRow {
                node_type: "tableRow".to_owned(),
                formatting: None,
                property_changes: None,
                structural_change: None,
                cells: vec![TableCell {
                    node_type: "tableCell".to_owned(),
                    formatting: Some(TableCellFormatting {
                        grid_span: Some(2.0),
                        ..TableCellFormatting::default()
                    }),
                    property_changes: None,
                    structural_change: None,
                    content: Vec::new(),
                }],
            }],
        };
        assert_eq!(
            serialize_table(&table, &mut context()).unwrap(),
            "<w:tbl><w:tblPr><w:tblStyle w:val=\"bad&quot;/&gt;&lt;evil&amp;\"/><w:tblW w:w=\"1008\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"1008\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"
        );
    }

    #[test]
    fn cell_merge_omits_absent_merge_attributes_and_only_first_change_emits() {
        let info = TrackedChangeInfo {
            id: -1.0,
            author: " ".to_owned(),
            date: None,
        };
        let xml = serialize_table_cell_formatting(
            None,
            None,
            Some(&TableStructuralChangeInfo {
                node_type: "tableCellMerge".to_owned(),
                info,
                v_merge: None,
                v_merge_orig: None,
            }),
        );
        assert_eq!(
            xml,
            "<w:tcPr><w:cellMerge w:id=\"0\" w:author=\"Unknown\"/></w:tcPr>"
        );
    }
}
