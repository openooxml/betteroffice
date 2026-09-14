//! Section-property leaves used by paragraph markers and document-body assembly.

use serde::{Deserialize, Serialize};

use crate::borders::{BorderSpec, parse_border_spec};
use crate::notes::{NoteProperties, parse_endnote_properties, parse_footnote_properties};
use crate::scalars::ColorValue;
use crate::xml::{XmlElement, parse_twips_measure};

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct Column {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct StoryReference {
    #[serde(rename = "type")]
    pub reference_type: String,
    #[serde(rename = "rId")]
    pub relationship_id: String,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LineNumbering {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub count_by: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub restart: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct PageBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
    #[serde(rename = "offsetFrom", skip_serializing_if = "Option::is_none")]
    pub offset_from: Option<String>,
    #[serde(rename = "zOrder", skip_serializing_if = "Option::is_none")]
    pub z_order: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionBackground {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentGrid {
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub grid_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_pitch: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub char_space: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageNumberingProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chapter_style: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chapter_separator: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gutter: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_count: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_space: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub equal_width: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub separator: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<Vec<Column>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_columns: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_start: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_references: Option<Vec<StoryReference>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_references: Option<Vec<StoryReference>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_pg: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_and_odd_headers: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_numbers: Option<LineNumbering>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<PageNumberingProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<PageBorders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<SectionBackground>,
    #[serde(rename = "footnotePr", skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<NoteProperties>,
    #[serde(rename = "endnotePr", skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<NoteProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid: Option<DocumentGrid>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paper_src_first: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paper_src_other: Option<f64>,
}

pub fn parse_section_properties(element: Option<&XmlElement>) -> SectionProperties {
    let Some(element) = element else {
        return SectionProperties::default();
    };
    let mut value = SectionProperties::default();
    if let Some(size) = element.child("w", "pgSz") {
        value.page_width = twips(size, "w", false);
        value.page_height = twips(size, "h", false);
        value.orientation = enum_attribute(size, "orient", &["landscape", "portrait"]);
    }
    if let Some(margins) = element.child("w", "pgMar") {
        value.margin_top = twips(margins, "top", true);
        value.margin_bottom = twips(margins, "bottom", true);
        value.margin_left = twips(margins, "left", false);
        value.margin_right = twips(margins, "right", false);
        value.header_distance = twips(margins, "header", false);
        value.footer_distance = twips(margins, "footer", false);
        value.gutter = twips(margins, "gutter", false);
    }
    if let Some(columns) = element.child("w", "cols") {
        value.column_count = numeric(columns, "num");
        value.column_space = twips(columns, "space", false);
        value.equal_width = explicit_bool_attribute(columns, "equalWidth");
        value.separator = truthy_attribute(columns, "sep");
        let definitions: Vec<_> = columns
            .children_named("w", "col")
            .map(|column| Column {
                width: twips(column, "w", false),
                space: twips(column, "space", false),
            })
            .collect();
        if !definitions.is_empty() {
            if value.column_count.is_none() {
                value.column_count = Some(definitions.len() as f64);
            }
            value.columns = Some(definitions);
        }
    }
    value.footnote_columns = element
        .child_by_local_name("footnoteColumns")
        .and_then(|child| numeric(child, "val"))
        .filter(|value| *value > 0.0);
    value.section_start = child_enum(
        element,
        "type",
        &[
            "continuous",
            "nextPage",
            "oddPage",
            "evenPage",
            "nextColumn",
        ],
    );
    value.vertical_align = child_enum(element, "vAlign", &["top", "center", "both", "bottom"]);
    value.bidi = boolean_child(element, "bidi");
    value.header_references = story_references(element, "headerReference");
    value.footer_references = story_references(element, "footerReference");
    value.title_pg = boolean_child(element, "titlePg");
    value.even_and_odd_headers = boolean_child(element, "evenAndOddHeaders");
    value.line_numbers = element.child("w", "lnNumType").map(|line| LineNumbering {
        start: numeric(line, "start"),
        count_by: numeric(line, "countBy"),
        distance: twips(line, "distance", false),
        restart: enum_attribute(line, "restart", &["continuous", "newPage", "newSection"]),
    });
    value.page_numbering =
        element
            .child("w", "pgNumType")
            .map(|numbering| PageNumberingProperties {
                start: numeric(numbering, "start"),
                format: numbering.attribute(Some("w"), "fmt").map(str::to_owned),
                chapter_style: numeric(numbering, "chapStyle"),
                chapter_separator: numbering.attribute(Some("w"), "chapSep").map(str::to_owned),
            });
    value.page_borders = parse_page_borders(element.child("w", "pgBorders"));
    value.background = parse_background(element.child("w", "background"));
    if let Some(properties) = element.child("w", "footnotePr") {
        let parsed = parse_footnote_properties(Some(properties));
        if parsed != NoteProperties::default() {
            value.footnote_pr = Some(parsed);
        }
    }
    if let Some(properties) = element.child("w", "endnotePr") {
        let parsed = parse_endnote_properties(Some(properties));
        if parsed != NoteProperties::default() {
            value.endnote_pr = Some(parsed);
        }
    }
    value.doc_grid = element.child("w", "docGrid").map(|grid| DocumentGrid {
        grid_type: enum_attribute(
            grid,
            "type",
            &["default", "lines", "linesAndChars", "snapToChars"],
        ),
        line_pitch: numeric(grid, "linePitch"),
        char_space: numeric(grid, "charSpace"),
    });
    if let Some(source) = element.child("w", "paperSrc") {
        value.paper_src_first = numeric(source, "first");
        value.paper_src_other = numeric(source, "other");
    }
    value
}

pub fn default_section_properties() -> SectionProperties {
    SectionProperties {
        page_width: Some(12_240.0),
        page_height: Some(15_840.0),
        orientation: Some("portrait".to_owned()),
        margin_top: Some(1_440.0),
        margin_bottom: Some(1_440.0),
        margin_left: Some(1_440.0),
        margin_right: Some(1_440.0),
        header_distance: Some(720.0),
        footer_distance: Some(720.0),
        gutter: Some(0.0),
        column_count: Some(1.0),
        column_space: Some(720.0),
        equal_width: Some(true),
        section_start: Some("nextPage".to_owned()),
        vertical_align: Some("top".to_owned()),
        ..SectionProperties::default()
    }
}

pub fn apply_section_inheritance(sections: &mut [SectionProperties]) {
    for index in 1..sections.len() {
        let previous = sections[index - 1].clone();
        inherit_references(
            &mut sections[index].header_references,
            previous.header_references.as_deref(),
        );
        inherit_references(
            &mut sections[index].footer_references,
            previous.footer_references.as_deref(),
        );
    }
}

fn story_references(parent: &XmlElement, local: &str) -> Option<Vec<StoryReference>> {
    let references: Vec<_> = parent
        .children_named("w", local)
        .map(|reference| StoryReference {
            reference_type: match reference.attribute(Some("w"), "type") {
                Some("first") => "first",
                Some("even") => "even",
                _ => "default",
            }
            .to_owned(),
            relationship_id: reference
                .attribute(Some("r"), "id")
                .unwrap_or_default()
                .to_owned(),
        })
        .collect();
    (!references.is_empty()).then_some(references)
}

fn inherit_references(own: &mut Option<Vec<StoryReference>>, prior: Option<&[StoryReference]>) {
    let Some(prior) = prior else { return };
    let own_types: Vec<_> = own
        .as_deref()
        .unwrap_or_default()
        .iter()
        .map(|reference| reference.reference_type.as_str())
        .collect();
    let inherited: Vec<_> = prior
        .iter()
        .filter(|reference| !own_types.contains(&reference.reference_type.as_str()))
        .cloned()
        .collect();
    if !inherited.is_empty() {
        own.get_or_insert_with(Vec::new).extend(inherited);
    }
}

fn parse_page_borders(element: Option<&XmlElement>) -> Option<PageBorders> {
    let element = element?;
    Some(PageBorders {
        top: parse_border_spec(element.child("w", "top")),
        bottom: parse_border_spec(element.child("w", "bottom")),
        left: parse_border_spec(element.child("w", "left")),
        right: parse_border_spec(element.child("w", "right")),
        display: enum_attribute(
            element,
            "display",
            &["allPages", "firstPage", "notFirstPage"],
        ),
        offset_from: enum_attribute(element, "offsetFrom", &["page", "text"]),
        z_order: enum_attribute(element, "zOrder", &["front", "back"]),
    })
}

fn parse_background(element: Option<&XmlElement>) -> Option<SectionBackground> {
    let element = element?;
    let rgb = element
        .attribute(Some("w"), "color")
        .filter(|value| *value != "auto")
        .map(|value| ColorValue {
            rgb: Some(value.to_owned()),
            ..ColorValue::default()
        });
    Some(SectionBackground {
        color: rgb,
        theme_color: element
            .attribute(Some("w"), "themeColor")
            .map(str::to_owned),
        theme_tint: element.attribute(Some("w"), "themeTint").map(str::to_owned),
        theme_shade: element
            .attribute(Some("w"), "themeShade")
            .map(str::to_owned),
    })
}

fn child_enum(parent: &XmlElement, child: &str, allowed: &[&str]) -> Option<String> {
    parent
        .child("w", child)
        .and_then(|element| enum_attribute(element, "val", allowed))
}

fn enum_attribute(element: &XmlElement, name: &str, allowed: &[&str]) -> Option<String> {
    element
        .attribute(Some("w"), name)
        .filter(|value| allowed.contains(value))
        .map(str::to_owned)
}

fn numeric(element: &XmlElement, name: &str) -> Option<f64> {
    element.parse_numeric_attribute(Some("w"), name, 1.0)
}

fn twips(element: &XmlElement, name: &str, signed: bool) -> Option<f64> {
    parse_twips_measure(element.attribute(Some("w"), name)?, signed)
}

fn boolean_child(parent: &XmlElement, child: &str) -> Option<bool> {
    parent
        .child("w", child)
        .map(|element| element.parse_boolean("w"))
}

fn explicit_bool_attribute(element: &XmlElement, name: &str) -> Option<bool> {
    match element.attribute(Some("w"), name) {
        Some("1" | "true") => Some(true),
        Some("0" | "false") => Some(false),
        _ => None,
    }
}

fn truthy_attribute(element: &XmlElement, name: &str) -> Option<bool> {
    matches!(element.attribute(Some("w"), name), Some("1" | "true")).then_some(true)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    #[test]
    fn universal_measures_preserve_section_geometry_and_round_trip() {
        let xml = br#"<w:sectPr><w:pgSz w:w="21cm" w:h="297mm"/><w:pgMar w:top="-1.25cm" w:bottom="-0.5in" w:left="10mm" w:right="36pt" w:header="3pc" w:footer="3pi" w:gutter="2.54mm"/><w:cols w:num="2" w:space="1.27cm"><w:col w:w="72pt" w:space="0mm"/></w:cols><w:lnNumType w:start="1" w:countBy="2" w:distance="0.25in"/><w:docGrid w:linePitch="360"/></w:sectPr>"#;
        let limits = ParseLimits::default();
        let document = parse_xml(xml, "word/document.xml", &mut ParseBudget::new(&limits)).unwrap();
        let properties = parse_section_properties(document.root());
        assert_eq!(properties.page_width, Some(11906.0));
        assert_eq!(properties.page_height, Some(16838.0));
        assert_eq!(properties.margin_top, Some(-709.0));
        assert_eq!(properties.margin_bottom, Some(-720.0));
        assert_eq!(properties.margin_left, Some(567.0));
        assert_eq!(properties.margin_right, Some(720.0));
        assert_eq!(properties.header_distance, Some(720.0));
        assert_eq!(properties.footer_distance, Some(720.0));
        assert_eq!(properties.gutter, Some(144.0));
        assert_eq!(properties.column_count, Some(2.0));
        assert_eq!(properties.column_space, Some(720.0));
        let column = &properties.columns.as_ref().unwrap()[0];
        assert_eq!(column.width, Some(1440.0));
        assert_eq!(column.space, Some(0.0));
        let numbering = properties.line_numbers.as_ref().unwrap();
        assert_eq!(numbering.start, Some(1.0));
        assert_eq!(numbering.count_by, Some(2.0));
        assert_eq!(numbering.distance, Some(360.0));
        assert_eq!(
            properties.doc_grid.as_ref().unwrap().line_pitch,
            Some(360.0)
        );
        let saved = crate::serializer::serialize_section_properties(Some(&properties));
        let reopened = parse_xml(
            saved.as_bytes(),
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(parse_section_properties(reopened.root()), properties);
    }

    #[test]
    fn malformed_section_measures_do_not_become_integer_prefixes() {
        let limits = ParseLimits::default();
        let document = parse_xml(
            br#"<w:sectPr><w:pgSz w:w="21CM" w:h="297px"/><w:pgMar w:top="10junk" w:left="-10mm" w:right="-720" w:bottom="10mmjunk" w:header="1e2pt" w:footer="+1pt" w:gutter="0.5"/><w:cols w:num="2suffix" w:space="5cmjunk"/><w:docGrid w:linePitch="360suffix"/></w:sectPr>"#,
            "word/document.xml", &mut ParseBudget::new(&limits)).unwrap();
        let properties = parse_section_properties(document.root());
        assert_eq!(
            properties,
            SectionProperties {
                column_count: Some(2.0),
                doc_grid: Some(DocumentGrid {
                    line_pitch: Some(360.0),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }

    #[test]
    fn parses_inline_section_markers_and_inherits_story_references() {
        let limits = ParseLimits::default();
        let document = parse_xml(
            br#"<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/><w:cols><w:col w:w="5000"/><w:col w:w="5000"/></w:cols><w15:footnoteColumns w:val="2"/><w:titlePg/></w:sectPr>"#,
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let first = parse_section_properties(document.root());
        assert_eq!(first.column_count, Some(2.0));
        assert_eq!(first.footnote_columns, Some(2.0));
        let mut sections = vec![first, SectionProperties::default()];
        apply_section_inheritance(&mut sections);
        assert_eq!(sections[1].title_pg, None);
        assert_eq!(
            sections[1].header_references.as_ref().unwrap()[0].relationship_id,
            "rId1"
        );
    }

    #[test]
    fn next_column_section_start_parses_and_round_trips() {
        let limits = ParseLimits::default();
        let document = parse_xml(
            br#"<w:sectPr><w:type w:val="nextColumn"/><w:cols w:num="2" w:space="720"/></w:sectPr>"#,
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let properties = parse_section_properties(document.root());
        assert_eq!(properties.section_start.as_deref(), Some("nextColumn"));
        assert!(
            crate::serializer::serialize_section_properties(Some(&properties))
                .contains(r#"<w:type w:val="nextColumn"/>"#)
        );
    }

    #[test]
    fn a_count_less_cols_with_authored_space_round_trips() {
        let limits = ParseLimits::default();
        let document = parse_xml(
            br#"<w:sectPr><w:cols w:space="708"/></w:sectPr>"#,
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let properties = parse_section_properties(document.root());
        assert_eq!(properties.column_space, Some(708.0));
        assert!(
            crate::serializer::serialize_section_properties(Some(&properties))
                .contains(r#"<w:cols w:space="708">"#)
        );
    }

    #[test]
    fn page_numbering_parses_and_round_trips() {
        let limits = ParseLimits::default();
        let document = parse_xml(
            br#"<w:sectPr><w:pgNumType w:fmt="lowerRoman" w:start="1" w:chapStyle="1" w:chapSep="hyphen"/></w:sectPr>"#,
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        let properties = parse_section_properties(document.root());
        let numbering = properties.page_numbering.as_ref().unwrap();
        assert_eq!(numbering.start, Some(1.0));
        assert_eq!(numbering.format.as_deref(), Some("lowerRoman"));
        assert_eq!(numbering.chapter_style, Some(1.0));
        assert_eq!(numbering.chapter_separator.as_deref(), Some("hyphen"));
        assert!(
            crate::serializer::serialize_section_properties(Some(&properties)).contains(
                r#"<w:pgNumType w:fmt="lowerRoman" w:start="1" w:chapStyle="1" w:chapSep="hyphen"/>"#
            )
        );
    }
}
