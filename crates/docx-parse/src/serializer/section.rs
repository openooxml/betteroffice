//! S10 `w:sectPr` serializer with schema and incumbent byte ordering.

use crate::notes::NoteProperties;
use crate::section::{DocumentGrid, PageBorders, SectionProperties, StoryReference};

use super::foundation::{BorderSide, write_border};
use super::xml_writer::{XmlWriter, int_attr, js_number};

/// Serialize final or paragraph-level section properties.
pub fn serialize_section_properties(properties: Option<&SectionProperties>) -> String {
    let Some(properties) = properties else {
        return String::new();
    };
    if !has_serialized_content(properties) {
        return String::new();
    }

    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:sectPr");

    for reference in properties.header_references.as_deref().unwrap_or_default() {
        write_story_reference(&mut writer, reference, true);
    }
    for reference in properties.footer_references.as_deref().unwrap_or_default() {
        write_story_reference(&mut writer, reference, false);
    }
    write_note_properties(&mut writer, properties.footnote_pr.as_ref(), true);
    write_note_properties(&mut writer, properties.endnote_pr.as_ref(), false);

    if let Some(value) = nonempty(properties.section_start.as_deref()) {
        writer
            .start_element("w:type")
            .attribute("w:val", value)
            .end_element();
    }
    write_page_size(&mut writer, properties);
    write_page_margins(&mut writer, properties);
    write_paper_source(&mut writer, properties);
    write_page_borders(&mut writer, properties.page_borders.as_ref());
    write_line_numbers(&mut writer, properties);
    write_columns(&mut writer, properties);

    if properties.footnote_columns.is_some_and(|value| value > 1.0) {
        writer
            .start_element("w15:footnoteColumns")
            .attribute("w:val", &int_attr(properties.footnote_columns))
            .end_element();
    }
    if let Some(value) = nonempty(properties.vertical_align.as_deref()) {
        writer
            .start_element("w:vAlign")
            .attribute("w:val", value)
            .end_element();
    }
    if properties.title_pg == Some(true) {
        writer.start_element("w:titlePg").end_element();
    }
    if properties.bidi == Some(true) {
        writer.start_element("w:bidi").end_element();
    }
    write_document_grid(&mut writer, properties.doc_grid.as_ref());
    if properties.even_and_odd_headers == Some(true) {
        writer.start_element("w:evenAndOddHeaders").end_element();
    }

    writer.end_element();
    writer.finish()
}

fn nonempty(value: Option<&str>) -> Option<&str> {
    value.filter(|value| !value.is_empty())
}

fn note_has_content(properties: Option<&NoteProperties>) -> bool {
    properties.is_some_and(|properties| {
        nonempty(properties.position.as_deref()).is_some()
            || nonempty(properties.num_fmt.as_deref()).is_some()
            || properties.num_start.is_some()
            || nonempty(properties.num_restart.as_deref()).is_some()
    })
}

fn page_size_has_content(properties: &SectionProperties) -> bool {
    properties.page_width.is_some()
        || properties.page_height.is_some()
        || properties.orientation.as_deref() == Some("landscape")
}

fn page_margins_have_content(properties: &SectionProperties) -> bool {
    properties.margin_top.is_some()
        || properties.margin_right.is_some()
        || properties.margin_bottom.is_some()
        || properties.margin_left.is_some()
        || properties.header_distance.is_some()
        || properties.footer_distance.is_some()
        || properties.gutter.is_some()
}

fn columns_have_content(properties: &SectionProperties) -> bool {
    let has_columns = properties
        .columns
        .as_ref()
        .is_some_and(|columns| !columns.is_empty());
    if properties.column_count.unwrap_or(0.0) == 0.0 && !has_columns {
        return false;
    }
    properties.column_count.is_some_and(|value| value > 1.0)
        || properties.column_space.is_some()
        || properties.equal_width.is_some()
        || properties.separator == Some(true)
        || has_columns
}

fn line_numbers_have_content(properties: &SectionProperties) -> bool {
    properties.line_numbers.as_ref().is_some_and(|line| {
        line.count_by.is_some()
            || line.start.is_some()
            || line.distance.is_some()
            || nonempty(line.restart.as_deref()).is_some()
    })
}

fn page_borders_have_content(properties: Option<&PageBorders>) -> bool {
    properties.is_some_and(|borders| {
        borders.top.is_some()
            || borders.left.is_some()
            || borders.bottom.is_some()
            || borders.right.is_some()
    })
}

fn document_grid_has_content(grid: Option<&DocumentGrid>) -> bool {
    grid.is_some_and(|grid| {
        nonempty(grid.grid_type.as_deref()).is_some()
            || grid.line_pitch.is_some()
            || grid.char_space.is_some()
    })
}

fn has_serialized_content(properties: &SectionProperties) -> bool {
    properties
        .header_references
        .as_ref()
        .is_some_and(|references| !references.is_empty())
        || properties
            .footer_references
            .as_ref()
            .is_some_and(|references| !references.is_empty())
        || note_has_content(properties.footnote_pr.as_ref())
        || note_has_content(properties.endnote_pr.as_ref())
        || nonempty(properties.section_start.as_deref()).is_some()
        || page_size_has_content(properties)
        || page_margins_have_content(properties)
        || properties.paper_src_first.is_some()
        || properties.paper_src_other.is_some()
        || page_borders_have_content(properties.page_borders.as_ref())
        || line_numbers_have_content(properties)
        || columns_have_content(properties)
        || properties.footnote_columns.is_some_and(|value| value > 1.0)
        || nonempty(properties.vertical_align.as_deref()).is_some()
        || properties.title_pg == Some(true)
        || properties.bidi == Some(true)
        || document_grid_has_content(properties.doc_grid.as_ref())
        || properties.even_and_odd_headers == Some(true)
}

fn write_story_reference(writer: &mut XmlWriter, reference: &StoryReference, header: bool) {
    writer
        .start_element(if header {
            "w:headerReference"
        } else {
            "w:footerReference"
        })
        .attribute("w:type", &reference.reference_type)
        .attribute("r:id", &reference.relationship_id)
        .end_element();
}

fn write_note_properties(
    writer: &mut XmlWriter,
    properties: Option<&NoteProperties>,
    footnote: bool,
) {
    let Some(properties) = properties.filter(|properties| note_has_content(Some(properties)))
    else {
        return;
    };
    writer.start_element(if footnote {
        "w:footnotePr"
    } else {
        "w:endnotePr"
    });
    if let Some(value) = nonempty(properties.position.as_deref()) {
        writer
            .start_element("w:pos")
            .attribute("w:val", value)
            .end_element();
    }
    if let Some(value) = nonempty(properties.num_fmt.as_deref()) {
        writer
            .start_element("w:numFmt")
            .attribute("w:val", value)
            .end_element();
    }
    if let Some(value) = properties.num_start {
        writer
            .start_element("w:numStart")
            .attribute("w:val", &js_number(value))
            .end_element();
    }
    if let Some(value) = nonempty(properties.num_restart.as_deref()) {
        writer
            .start_element("w:numRestart")
            .attribute("w:val", value)
            .end_element();
    }
    writer.end_element();
}

fn write_page_size(writer: &mut XmlWriter, properties: &SectionProperties) {
    if !page_size_has_content(properties) {
        return;
    }
    writer.start_element("w:pgSz");
    if let Some(value) = properties.page_width {
        writer.attribute("w:w", &int_attr(Some(value)));
    }
    if let Some(value) = properties.page_height {
        writer.attribute("w:h", &int_attr(Some(value)));
    }
    if properties.orientation.as_deref() == Some("landscape") {
        writer.attribute("w:orient", "landscape");
    }
    writer.end_element();
}

fn write_page_margins(writer: &mut XmlWriter, properties: &SectionProperties) {
    if !page_margins_have_content(properties) {
        return;
    }
    writer.start_element("w:pgMar");
    for (name, value) in [
        ("w:top", properties.margin_top),
        ("w:right", properties.margin_right),
        ("w:bottom", properties.margin_bottom),
        ("w:left", properties.margin_left),
        ("w:header", properties.header_distance),
        ("w:footer", properties.footer_distance),
        ("w:gutter", properties.gutter),
    ] {
        if let Some(value) = value {
            writer.attribute(name, &int_attr(Some(value)));
        }
    }
    writer.end_element();
}

fn write_paper_source(writer: &mut XmlWriter, properties: &SectionProperties) {
    if properties.paper_src_first.is_none() && properties.paper_src_other.is_none() {
        return;
    }
    writer.start_element("w:paperSrc");
    if let Some(value) = properties.paper_src_first {
        writer.attribute("w:first", &js_number(value));
    }
    if let Some(value) = properties.paper_src_other {
        writer.attribute("w:other", &js_number(value));
    }
    writer.end_element();
}

fn write_page_borders(writer: &mut XmlWriter, borders: Option<&PageBorders>) {
    let Some(borders) = borders.filter(|borders| page_borders_have_content(Some(borders))) else {
        return;
    };
    writer.start_element("w:pgBorders");
    if let Some(value) = nonempty(borders.display.as_deref()) {
        writer.attribute("w:display", value);
    }
    if let Some(value) = nonempty(borders.offset_from.as_deref()) {
        writer.attribute("w:offsetFrom", value);
    }
    if let Some(value) = nonempty(borders.z_order.as_deref()) {
        writer.attribute("w:zOrder", value);
    }
    for (border, side) in [
        (borders.top.as_ref(), BorderSide::Top),
        (borders.left.as_ref(), BorderSide::Left),
        (borders.bottom.as_ref(), BorderSide::Bottom),
        (borders.right.as_ref(), BorderSide::Right),
    ] {
        if let Some(border) = border {
            write_border(writer, border, side);
        }
    }
    writer.end_element();
}

fn write_line_numbers(writer: &mut XmlWriter, properties: &SectionProperties) {
    let Some(line) = properties
        .line_numbers
        .as_ref()
        .filter(|_| line_numbers_have_content(properties))
    else {
        return;
    };
    writer.start_element("w:lnNumType");
    for (name, value) in [
        ("w:countBy", line.count_by),
        ("w:start", line.start),
        ("w:distance", line.distance),
    ] {
        if let Some(value) = value {
            writer.attribute(name, &int_attr(Some(value)));
        }
    }
    if let Some(value) = nonempty(line.restart.as_deref()) {
        writer.attribute("w:restart", value);
    }
    writer.end_element();
}

fn write_columns(writer: &mut XmlWriter, properties: &SectionProperties) {
    if !columns_have_content(properties) {
        return;
    }
    writer.start_element("w:cols");
    if properties.column_count.is_some_and(|value| value > 1.0) {
        writer.attribute("w:num", &int_attr(properties.column_count));
    }
    if let Some(value) = properties.column_space {
        writer.attribute("w:space", &int_attr(Some(value)));
    }
    if let Some(value) = properties.equal_width {
        writer.attribute("w:equalWidth", if value { "1" } else { "0" });
    }
    if properties.separator == Some(true) {
        writer.attribute("w:sep", "1");
    }
    let columns = properties.columns.as_deref().unwrap_or_default();
    if columns.is_empty() {
        // The incumbent template spells w:cols as an explicit start/end pair,
        // even when it only contains attributes.
        writer.text("");
    }
    for column in columns {
        writer.start_element("w:col");
        if let Some(value) = column.width {
            writer.attribute("w:w", &int_attr(Some(value)));
        }
        if let Some(value) = column.space {
            writer.attribute("w:space", &int_attr(Some(value)));
        }
        if column.width.is_none() && column.space.is_none() {
            writer.end_empty_element_with_space();
        } else {
            writer.end_element();
        }
    }
    writer.end_element();
}

fn write_document_grid(writer: &mut XmlWriter, grid: Option<&DocumentGrid>) {
    let Some(grid) = grid.filter(|grid| document_grid_has_content(Some(grid))) else {
        return;
    };
    writer.start_element("w:docGrid");
    if let Some(value) = nonempty(grid.grid_type.as_deref()) {
        writer.attribute("w:type", value);
    }
    if let Some(value) = grid.line_pitch {
        writer.attribute("w:linePitch", &js_number(value));
    }
    if let Some(value) = grid.char_space {
        writer.attribute("w:charSpace", &js_number(value));
    }
    writer.end_element();
}

#[cfg(test)]
mod tests {
    use crate::borders::BorderSpec;
    use crate::notes::NoteProperties;
    use crate::section::{Column, DocumentGrid, LineNumbering, PageBorders, StoryReference};

    use super::*;

    #[test]
    fn omits_absent_and_empty_section_properties() {
        assert_eq!(serialize_section_properties(None), "");
        assert_eq!(
            serialize_section_properties(Some(&SectionProperties::default())),
            ""
        );
        let properties = SectionProperties {
            columns: Some(Vec::new()),
            line_numbers: Some(LineNumbering::default()),
            doc_grid: Some(DocumentGrid::default()),
            page_borders: Some(PageBorders::default()),
            ..SectionProperties::default()
        };
        assert_eq!(serialize_section_properties(Some(&properties)), "");
    }

    #[test]
    fn preserves_schema_order_integer_rules_and_empty_column_spelling() {
        let properties = SectionProperties {
            header_references: Some(vec![StoryReference {
                reference_type: "default".to_owned(),
                relationship_id: "rId1".to_owned(),
            }]),
            footnote_pr: Some(NoteProperties {
                position: Some("pageBottom".to_owned()),
                num_start: Some(1.25),
                ..NoteProperties::default()
            }),
            section_start: Some("nextPage".to_owned()),
            page_width: Some(1008.000_000_000_000_1),
            page_height: Some(2000.5),
            orientation: Some("landscape".to_owned()),
            margin_top: Some(-1.5),
            paper_src_first: Some(2.5),
            line_numbers: Some(LineNumbering {
                count_by: Some(4.5),
                ..LineNumbering::default()
            }),
            column_count: Some(2.0),
            column_space: Some(720.0),
            columns: Some(vec![Column::default()]),
            footnote_columns: Some(3.0),
            vertical_align: Some("center".to_owned()),
            title_pg: Some(true),
            bidi: Some(true),
            doc_grid: Some(DocumentGrid {
                grid_type: Some("lines".to_owned()),
                line_pitch: Some(360.25),
                char_space: None,
            }),
            even_and_odd_headers: Some(true),
            ..SectionProperties::default()
        };
        assert_eq!(
            serialize_section_properties(Some(&properties)),
            "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId1\"/><w:footnotePr><w:pos w:val=\"pageBottom\"/><w:numStart w:val=\"1.25\"/></w:footnotePr><w:type w:val=\"nextPage\"/><w:pgSz w:w=\"1008\" w:h=\"2001\" w:orient=\"landscape\"/><w:pgMar w:top=\"-1\"/><w:paperSrc w:first=\"2.5\"/><w:lnNumType w:countBy=\"5\"/><w:cols w:num=\"2\" w:space=\"720\"><w:col /></w:cols><w15:footnoteColumns w:val=\"3\"/><w:vAlign w:val=\"center\"/><w:titlePg/><w:bidi/><w:docGrid w:type=\"lines\" w:linePitch=\"360.25\"/><w:evenAndOddHeaders/></w:sectPr>"
        );
        assert_eq!(
            serialize_section_properties(Some(&SectionProperties {
                column_count: Some(2.0),
                column_space: Some(720.0),
                ..SectionProperties::default()
            })),
            "<w:sectPr><w:cols w:num=\"2\" w:space=\"720\"></w:cols></w:sectPr>"
        );
    }

    #[test]
    fn escapes_every_attacker_derived_section_string() {
        let attack = "\"/><evil attr='&".to_owned();
        let properties = SectionProperties {
            header_references: Some(vec![StoryReference {
                reference_type: attack.clone(),
                relationship_id: attack.clone(),
            }]),
            endnote_pr: Some(NoteProperties {
                position: Some(attack.clone()),
                num_fmt: Some(attack.clone()),
                num_restart: Some(attack.clone()),
                ..NoteProperties::default()
            }),
            section_start: Some(attack.clone()),
            page_borders: Some(PageBorders {
                top: Some(BorderSpec {
                    style: attack.clone(),
                    color: None,
                    size: None,
                    space: None,
                    shadow: None,
                    frame: None,
                }),
                display: Some(attack.clone()),
                offset_from: Some(attack.clone()),
                z_order: Some(attack.clone()),
                ..PageBorders::default()
            }),
            line_numbers: Some(LineNumbering {
                restart: Some(attack.clone()),
                ..LineNumbering::default()
            }),
            vertical_align: Some(attack.clone()),
            doc_grid: Some(DocumentGrid {
                grid_type: Some(attack),
                ..DocumentGrid::default()
            }),
            ..SectionProperties::default()
        };
        let xml = serialize_section_properties(Some(&properties));
        assert!(!xml.contains("<evil"));
        assert!(!xml.contains("attr='"));
        assert_eq!(
            xml.matches("&quot;/&gt;&lt;evil attr=&apos;&amp;").count(),
            13
        );
    }
}
