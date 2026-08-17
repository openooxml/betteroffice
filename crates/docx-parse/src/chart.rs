//! DOCX packaging around the shared DrawingML chart part.

use indexmap::IndexMap;
use ooxml_drawingml::chart::{ChartXml, parse_chart_space};
use ooxml_drawingml::resolve_color_value_to_hex;
use serde::{Deserialize, Serialize};

use crate::drawingml::parse_color_element;
use crate::image::{
    ImagePosition, ImageSize, ImageWrap, PositionAxis, parse_position_h, parse_position_v,
    parse_wrap_element_without_distances,
};
use crate::relationships::{
    RelationshipMap, TargetMode, relationship_types, resolve_relative_path,
};
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

pub use ooxml_drawingml::chart::{
    ChartAxes, ChartAxis, ChartLegend, ChartMarker, ChartPlotGroup, ChartPoint, ChartSeries,
};

const MAX_DEEP_DEPTH: usize = 64;

pub type ChartPartsMap = IndexMap<String, Chart>;

impl ChartXml for XmlElement {
    fn local_name(&self) -> &str {
        XmlElement::local_name(self)
    }

    fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        XmlElement::attribute(self, prefix, name)
    }

    fn child_elements(&self) -> impl Iterator<Item = &Self> {
        XmlElement::child_elements(self)
    }

    fn descendant_text(&self) -> String {
        self.text_content()
    }

    fn solid_fill_hex(&self) -> Option<String> {
        resolve_color_value_to_hex(parse_color_element(Some(self)).as_ref())
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Chart {
    #[serde(rename = "type")]
    pub content_type: String,
    pub chart_type: String,
    #[serde(rename = "rId", skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend: Option<ChartLegend>,
    pub series: Vec<ChartSeries>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axes: Option<ChartAxes>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<ImageSize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<ImageWrap>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImagePosition>,
    pub plot_groups: Vec<ChartPlotGroup>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_list: Option<Vec<ChartAxis>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<f64>,
}

pub fn parse_chart_xml(
    xml: &[u8],
    path: Option<&str>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<Chart>, ParseError> {
    let document = parse_xml(xml, part, budget)?;
    let Some(chart_space) = document.root() else {
        return Ok(None);
    };
    let Some(space) = parse_chart_space(chart_space) else {
        return Ok(None);
    };
    Ok(Some(Chart {
        content_type: "chart".to_owned(),
        chart_type: space.chart_type,
        relationship_id: None,
        path: path.map(str::to_owned),
        title: space.title,
        legend: space.legend,
        series: space.series,
        axes: space.axes,
        size: None,
        wrap: None,
        position: None,
        plot_groups: space.plot_groups,
        axis_list: space.axis_list,
        description: None,
        decorative: None,
        relative_height: None,
    }))
}

/// Every chart the package carries, keyed by its normalized path and by the
/// `word/`-relative alias a drawing may reference it under.
///
/// A chart part is a decorative leaf: nothing structural is read from it and
/// nothing structural is read after it. So each gets its own [`ParseBudget`] —
/// one cached series can neither starve the body nor starve the next chart on
/// depth, attributes or text — and a part `parse_xml` refuses is skipped,
/// leaving the document to open with that chart missing exactly as it opens
/// when the part is absent. `parse_xml` reports only `MalformedXml`,
/// `UnsafeXml` and `ResourceLimit`, each naming this part; a hostile package is
/// refused earlier, by `ooxml_opc::unzip_parts`. A path
/// [`normalize_chart_path`] rejects cannot name a chart part either, so it is
/// skipped the same way.
///
/// What bounds the whole is one shared allowance of `max_xml_events`: every
/// part draws its event budget from the same remainder, so all the charts
/// together cost no more reads than one part was allowed before. A part the
/// remainder cannot cover is declined like a malformed one. Nothing else is
/// shared, so per-part isolation still holds.
pub fn parse_chart_parts(
    all_xml: &IndexMap<String, Vec<u8>>,
    limits: &ParseLimits,
) -> ChartPartsMap {
    let mut charts = ChartPartsMap::new();
    let mut events = limits.max_xml_events;
    for (path, xml) in all_xml {
        let Some(normalized) = chart_part_path(path) else {
            continue;
        };
        let part_limits = ParseLimits {
            max_xml_events: events,
            ..limits.clone()
        };
        let mut budget = ParseBudget::new(&part_limits);
        let parsed = parse_chart_xml(xml, Some(&normalized), path, &mut budget);
        events -= budget.xml_events();
        let Ok(Some(chart)) = parsed else {
            continue;
        };
        charts.insert(normalized.clone(), chart.clone());
        charts.insert(
            normalized
                .strip_prefix("word/")
                .unwrap_or(&normalized)
                .to_owned(),
            chart,
        );
    }
    charts
}

/// Whether [`parse_chart_parts`] reads `path` as a chart part.
pub fn is_chart_part(path: &str) -> bool {
    chart_part_path(path).is_some()
}

fn chart_part_path(path: &str) -> Option<String> {
    let normalized = normalize_chart_path(path).ok()?;
    let lower = normalized.to_ascii_lowercase();
    (lower.starts_with("word/charts/") && lower.ends_with(".xml")).then_some(normalized)
}

pub fn normalize_chart_path(target: &str) -> Result<String, ParseError> {
    if target.is_empty() {
        return Ok(String::new());
    }
    let mut normalized = target.trim_start_matches('/').to_owned();
    if normalized.starts_with("../") {
        // Pinned quirk: this branch is not subsequently prefixed with `word/`.
        normalized = resolve_relative_path("word/document.xml", &normalized)?;
    } else if normalized.starts_with("charts/") {
        normalized = format!("word/{normalized}");
    } else if !normalized.starts_with("word/") {
        normalized = format!("word/{normalized}");
    }
    Ok(normalized)
}

/// What chart reading made of one `<w:drawing>`.
#[derive(Clone, Debug, PartialEq)]
pub enum DrawingChart {
    /// The drawing names no internal chart part.
    None,
    /// The drawing names a chart part this package has no chart for: the
    /// parser declined it, or it is missing. Anything else the drawing holds
    /// belongs to that chart, so it must not be read as a picture.
    Unread,
    /// The chart the drawing names.
    Chart(Box<Chart>),
}

pub fn parse_chart_from_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    charts: Option<&ChartPartsMap>,
) -> Result<DrawingChart, ParseError> {
    let (Some(relationships), Some(charts)) = (relationships, charts) else {
        return Ok(DrawingChart::None);
    };
    let Some(chart_ref) = first_deep(drawing, "chart", 0) else {
        return Ok(DrawingChart::None);
    };
    let Some(relationship_id) = chart_ref.attribute(Some("r"), "id") else {
        return Ok(DrawingChart::None);
    };
    let Some(relationship) = relationships.get(relationship_id) else {
        return Ok(DrawingChart::None);
    };
    if relationship.relationship_type != relationship_types::CHART
        || relationship.target_mode == Some(TargetMode::External)
    {
        return Ok(DrawingChart::None);
    }
    let path = normalize_chart_path(&relationship.target)?;
    let alias = path.strip_prefix("word/").unwrap_or(&path);
    let Some(source) = charts.get(&path).or_else(|| charts.get(alias)) else {
        return Ok(DrawingChart::Unread);
    };
    let mut chart = source.clone();
    apply_drawing_metadata(&mut chart, drawing);
    chart.relationship_id = Some(relationship_id.to_owned());
    chart.path = Some(path);
    chart.size = parse_drawing_extent(drawing).or(chart.size);
    Ok(DrawingChart::Chart(Box::new(chart)))
}

fn first_deep<'a>(root: &'a XmlElement, local: &str, depth: usize) -> Option<&'a XmlElement> {
    if depth > MAX_DEEP_DEPTH {
        return None;
    }
    if root.local_name() == local {
        return Some(root);
    }
    root.child_elements()
        .find_map(|child| first_deep(child, local, depth + 1))
}

fn parse_number(raw: Option<&str>) -> Option<f64> {
    let value = raw?.trim();
    if value.is_empty() {
        return None;
    }
    let parsed = if let Some(hex) = value
        .strip_prefix("0x")
        .or_else(|| value.strip_prefix("0X"))
    {
        u64::from_str_radix(hex, 16).ok().map(|value| value as f64)
    } else if let Some(binary) = value
        .strip_prefix("0b")
        .or_else(|| value.strip_prefix("0B"))
    {
        u64::from_str_radix(binary, 2)
            .ok()
            .map(|value| value as f64)
    } else if let Some(octal) = value
        .strip_prefix("0o")
        .or_else(|| value.strip_prefix("0O"))
    {
        u64::from_str_radix(octal, 8).ok().map(|value| value as f64)
    } else {
        value.parse::<f64>().ok()
    }?;
    parsed.is_finite().then_some(parsed)
}

fn parse_drawing_extent(drawing: &XmlElement) -> Option<ImageSize> {
    let container = drawing
        .child_elements()
        .find(|child| matches!(child.name.as_str(), "wp:inline" | "wp:anchor"))?;
    let extent = container.child_by_full_name("wp:extent")?;
    let width = extent.parse_numeric_attribute(None, "cx", 1.0)?;
    let height = extent.parse_numeric_attribute(None, "cy", 1.0)?;
    Some(ImageSize { width, height })
}

fn apply_drawing_metadata(chart: &mut Chart, drawing: &XmlElement) {
    let Some(container) = drawing
        .child_elements()
        .find(|child| matches!(child.name.as_str(), "wp:inline" | "wp:anchor"))
    else {
        return;
    };
    let doc_pr = container.child_by_full_name("wp:docPr");
    chart.description = doc_pr
        .and_then(|value| value.attribute(None, "descr"))
        .map(str::to_owned);
    chart.decorative =
        (doc_pr.and_then(|value| value.attribute(None, "decorative")) == Some("1")).then_some(true);
    let hidden = doc_pr.and_then(|value| value.attribute(None, "hidden")) == Some("1");
    if container.name == "wp:inline" {
        chart.wrap = Some(ImageWrap {
            wrap_type: "inline".to_owned(),
            wrap_text: None,
            dist_t: None,
            dist_b: None,
            dist_l: None,
            dist_r: None,
            polygon: None,
        });
        return;
    }
    let behind_doc = container.attribute(None, "behindDoc") == Some("1");
    let wrap_element = container.child_elements().find(|child| {
        matches!(
            child.name.as_str(),
            "wp:wrapNone"
                | "wp:wrapSquare"
                | "wp:wrapTight"
                | "wp:wrapThrough"
                | "wp:wrapTopAndBottom"
        )
    });
    chart.wrap = Some(parse_wrap_element_without_distances(
        wrap_element,
        behind_doc,
    ));
    let relative_height = parse_number(container.attribute(None, "relativeHeight"));
    chart.relative_height = relative_height;
    chart.position = Some(ImagePosition {
        use_simple_pos: None,
        simple_pos: None,
        relative_height,
        behind_doc: behind_doc.then_some(true),
        hidden: hidden.then_some(true),
        locked: None,
        horizontal: parse_position_h(container.child_by_full_name("wp:positionH")).unwrap_or(
            PositionAxis {
                relative_to: "column".to_owned(),
                alignment: None,
                pos_offset: None,
                offset: None,
            },
        ),
        vertical: parse_position_v(container.child_by_full_name("wp:positionV")).unwrap_or(
            PositionAxis {
                relative_to: "paragraph".to_owned(),
                alignment: None,
                pos_offset: None,
                offset: None,
            },
        ),
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::block::BlockContent;
    use crate::document::get_paragraph_text;
    use crate::inline::{InlineNode, RunContent};
    use crate::paragraph::ParagraphContent;
    use crate::relationships::{Relationship, TargetMode};
    use crate::s4::parse_docx_s4_projection;
    use crate::s6::parse_docx_s6_projection;
    use crate::s8::parse_docx_s8_projection;
    use crate::s9::{S9ParseOptions, parse_docx_s9_wire, parse_docx_s9_wire_with_limits};
    use crate::xml::ParseLimits;

    fn parse(xml: &str) -> Chart {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_chart_xml(
            xml.as_bytes(),
            Some("word/charts/chart1.xml"),
            "chart1.xml",
            &mut budget,
        )
        .unwrap()
        .unwrap()
    }

    #[test]
    fn parses_combo_chart_with_pinned_explicit_defaults() {
        let chart = parse(
            r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t> Sales </a:t></a:r></a:p></c:rich></c:tx></c:title><c:legend><c:legendPos val="r"/></c:legend><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:tx><c:v>A</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/></c:barChart><c:lineChart/><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:delete val="0"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        );
        assert_eq!(chart.chart_type, "column");
        assert_eq!(chart.title.as_deref(), Some("Sales"));
        assert_eq!(chart.legend.unwrap().position.as_deref(), Some("right"));
        assert_eq!(chart.series[0].categories, ["Q1"]);
        assert_eq!(chart.series[0].values, [2.0]);
        assert_eq!(chart.plot_groups.len(), 2);
        assert!(!chart.axis_list.as_ref().unwrap()[0].reversed);
        assert!(chart.axis_list.as_ref().unwrap()[1].reversed);
        assert!(!chart.axis_list.as_ref().unwrap()[0].hidden);
    }

    #[test]
    fn preserves_normalization_branch_quirk_and_rejects_external_chart() {
        assert_eq!(
            normalize_chart_path("charts/a.xml").unwrap(),
            "word/charts/a.xml"
        );
        assert_eq!(
            normalize_chart_path("../charts/a.xml").unwrap(),
            "charts/a.xml"
        );
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let source = parse_chart_xml(
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart/></c:plotArea></c:chart></c:chartSpace>"#,
            Some("word/charts/a.xml"),
            "chart",
            &mut budget,
        )
        .unwrap()
        .unwrap();
        let mut charts = ChartPartsMap::new();
        charts.insert("word/charts/a.xml".to_owned(), source);
        let mut relationships = RelationshipMap::new();
        relationships.insert(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: relationship_types::CHART.to_owned(),
                target: "charts/a.xml".to_owned(),
                target_mode: Some(TargetMode::External),
            },
        );
        let document = crate::xml::parse_xml(
            br#"<w:drawing xmlns:w="w" xmlns:wp="wp" xmlns:c="c" xmlns:r="r"><wp:inline><wp:extent cx="1" cy="2"/><a:graphic xmlns:a="a"><c:chart r:id="rId1"/></a:graphic></wp:inline></w:drawing>"#,
            "drawing",
            &mut budget,
        )
        .unwrap();
        assert_eq!(
            parse_chart_from_drawing(
                document.root().unwrap(),
                Some(&relationships),
                Some(&charts)
            )
            .unwrap(),
            DrawingChart::None
        );
    }

    #[test]
    fn truncates_deep_search_and_points_without_panicking() {
        let nested = format!(
            "<c:chartSpace xmlns:c=\"c\">{}<c:plotArea><c:lineChart/></c:plotArea>{}</c:chartSpace>",
            "<x>".repeat(66),
            "</x>".repeat(66)
        );
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        assert!(
            parse_chart_xml(nested.as_bytes(), None, "deep-chart", &mut budget)
                .unwrap()
                .is_none()
        );
    }

    const CHART_PART: &str = "word/charts/chart1.xml";
    const MALFORMED: &[u8] = b"<c:chartSpace><c:chart></c:chartSpace>";

    const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/word/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#;
    const ROOT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    const DOCUMENT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/><Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/></Relationships>"#;
    const DOCUMENT_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p w14:paraId="11111111"><w:r><w:t>Body text</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:p w14:paraId="33333333"><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="2" name="Chart 2"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart2"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="44444444"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#;

    /// A chart part whose only weight is its cache. Every point costs the
    /// reader five events.
    fn chart_space(points: usize) -> Vec<u8> {
        let mut xml = String::from(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:val><c:numRef><c:numCache>"#,
        );
        for point in 0..points {
            xml.push_str(&format!("<c:pt><c:v>{point}</c:v></c:pt>"));
        }
        xml.push_str(
            "</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>",
        );
        xml.into_bytes()
    }

    /// A document of one paragraph, two charted paragraphs and a table.
    fn chart_docx() -> Vec<(String, Vec<u8>)> {
        vec![
            ("[Content_Types].xml".to_owned(), CONTENT_TYPES.to_vec()),
            ("_rels/.rels".to_owned(), ROOT_RELS.to_vec()),
            (
                "word/_rels/document.xml.rels".to_owned(),
                DOCUMENT_RELS.to_vec(),
            ),
            ("word/document.xml".to_owned(), DOCUMENT_XML.to_vec()),
            ("word/charts/chart1.xml".to_owned(), chart_space(1)),
            ("word/charts/chart2.xml".to_owned(), chart_space(1)),
        ]
    }

    /// [`chart_docx`] with `part`'s bytes replaced.
    fn chart_docx_with(part: &str, bytes: &[u8]) -> Vec<u8> {
        let mut parts = chart_docx();
        let slot = parts
            .iter_mut()
            .find(|(path, _)| path == part)
            .unwrap_or_else(|| panic!("{part} is not in the fixture"));
        slot.1 = bytes.to_vec();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// [`DOCUMENT_XML`] with `count` more one-run paragraphs before the section.
    fn padded_document_xml(count: usize) -> Vec<u8> {
        let padding = (0..count)
            .map(|index| {
                format!(
                    r#"<w:p w14:paraId="{:08X}"><w:r><w:t>pad</w:t></w:r></w:p>"#,
                    0x0100_0000 + index
                )
            })
            .collect::<String>();
        String::from_utf8(DOCUMENT_XML.to_vec())
            .unwrap()
            .replace("<w:sectPr>", &format!("{padding}<w:sectPr>"))
            .into_bytes()
    }

    /// A document of `count` charted paragraphs, each chart caching `points`.
    fn many_chart_docx(count: usize, points: usize) -> Vec<u8> {
        let mut overrides = String::new();
        let mut relationships = String::new();
        let mut body = String::new();
        let mut chart_parts = Vec::new();
        for index in 1..=count {
            overrides.push_str(&format!(
                r#"<Override PartName="/word/charts/chart{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#
            ));
            relationships.push_str(&format!(
                r#"<Relationship Id="rIdChart{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{index}.xml"/>"#
            ));
            body.push_str(&format!(
                r#"<w:p w14:paraId="{:08X}"><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="{index}" name="Chart {index}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart{index}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
                0x0200_0000 + index
            ));
            chart_parts.push((format!("word/charts/chart{index}.xml"), chart_space(points)));
        }
        let mut parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{overrides}</Types>"#
                )
                .into_bytes(),
            ),
            ("_rels/.rels".to_owned(), ROOT_RELS.to_vec()),
            (
                "word/_rels/document.xml.rels".to_owned(),
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
                )
                .into_bytes(),
            ),
            (
                "word/document.xml".to_owned(),
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
                )
                .into_bytes(),
            ),
        ];
        parts.extend(chart_parts);
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// The drawings in `content` no chart was read for.
    fn opaque_drawings(content: &[BlockContent]) -> usize {
        content
            .iter()
            .filter_map(|block| match block {
                BlockContent::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .flat_map(|paragraph| &paragraph.content)
            .filter_map(|item| match item {
                ParagraphContent::Inline(InlineNode::Run(run)) => Some(run),
                _ => None,
            })
            .flat_map(|run| &run.content)
            .filter(|item| matches!(item, RunContent::OpaqueDrawing { .. }))
            .count()
    }

    /// The chart part every charted run in `content` resolved to.
    fn body_charts(content: &[BlockContent]) -> Vec<String> {
        content
            .iter()
            .filter_map(|block| match block {
                BlockContent::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .flat_map(|paragraph| &paragraph.content)
            .filter_map(|item| match item {
                ParagraphContent::Inline(InlineNode::Run(run)) => Some(run),
                _ => None,
            })
            .flat_map(|run| &run.content)
            .filter_map(|item| match item {
                RunContent::Chart { chart } => chart.path.clone(),
                _ => None,
            })
            .collect()
    }

    /// Opens a document whose first chart part carries `bytes` and asserts the
    /// document is whole apart from that one chart.
    fn assert_damaged_chart_is_dropped(bytes: &[u8]) {
        let package = parse_docx_s9_wire(
            &chart_docx_with(CHART_PART, bytes),
            S9ParseOptions::default(),
        )
        .unwrap()
        .document
        .package;
        let content = &package.document.content;
        assert_eq!(content.len(), 4);
        let BlockContent::Paragraph(first) = &content[0] else {
            panic!("the first block is a paragraph");
        };
        assert_eq!(get_paragraph_text(first), "Body text");
        assert!(matches!(&content[3], BlockContent::Table(_)));
        assert_eq!(
            package
                .chart_entries
                .iter()
                .map(|(path, _)| path.as_str())
                .collect::<Vec<_>>(),
            ["word/charts/chart2.xml", "charts/chart2.xml"]
        );
        assert_eq!(body_charts(content), ["word/charts/chart2.xml"]);
        assert_eq!(opaque_drawings(content), 1);
    }

    #[test]
    fn a_malformed_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(MALFORMED);
    }

    #[test]
    fn an_empty_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(b"");
    }

    /// A chart part that reads but models nothing is declined the same way.
    #[test]
    fn a_chart_part_without_a_chart_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(br#"<c:chartSpace xmlns:c="c"/>"#);
    }

    #[test]
    fn a_doctype_bearing_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(
            br#"<?xml version="1.0"?><!DOCTYPE c:chartSpace [<!ENTITY x "y">]><c:chartSpace xmlns:c="c"/>"#,
        );
    }

    #[test]
    fn a_truncated_chart_part_drops_only_that_chart() {
        let source = chart_space(8);
        assert_damaged_chart_is_dropped(&source[..source.len() / 2]);
    }

    #[test]
    fn a_binary_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(&[0x00, 0xff, 0xfe, 0x01, 0x80, 0x7f]);
    }

    #[test]
    fn an_undeclared_entity_in_a_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:title>&nbsp;</c:title></c:chart></c:chartSpace>"#,
        );
    }

    /// Every stage that reads chart parts declines the same part.
    #[test]
    fn a_damaged_chart_part_stops_no_projection_stage() {
        let document = chart_docx_with(CHART_PART, MALFORMED);
        assert_eq!(
            parse_docx_s4_projection(&document)
                .unwrap()
                .chart_entries
                .iter()
                .map(|(path, _)| path.as_str())
                .collect::<Vec<_>>(),
            ["word/charts/chart2.xml", "charts/chart2.xml"]
        );
        for body in [
            parse_docx_s6_projection(&document).unwrap().body,
            parse_docx_s8_projection(&document).unwrap().body,
        ] {
            assert_eq!(body_charts(&body.content), ["word/charts/chart2.xml"]);
        }
    }

    /// The undamaged document that used to be refused: spec-valid charts whose
    /// caches fill the chart allowance on their own, in a body the same budget
    /// could not have covered too.
    #[test]
    fn a_document_whose_charts_fill_the_chart_budget_opens_with_every_chart() {
        let mut parts = chart_docx();
        for (path, bytes) in &mut parts {
            if path.starts_with("word/charts/") {
                *bytes = chart_space(4_000);
            }
            if path == "word/document.xml" {
                *bytes = padded_document_xml(400);
            }
        }
        let document = ooxml_opc::rezip_parts(&parts).unwrap();
        let limits = ParseLimits {
            max_xml_events: 41_000,
            ..ParseLimits::default()
        };

        let package = parse_docx_s9_wire_with_limits(&document, S9ParseOptions::default(), &limits)
            .unwrap()
            .document
            .package;

        assert_eq!(
            body_charts(&package.document.content),
            ["word/charts/chart1.xml", "word/charts/chart2.xml"]
        );
        assert!(
            package
                .chart_entries
                .iter()
                .all(|(_, chart)| chart.series[0].values.len() == 4_000)
        );
    }

    /// Chart parts draw on one shared event allowance. The parts it covers are
    /// read; the rest are declined and their drawings stay opaque.
    #[test]
    fn charts_past_the_shared_event_budget_are_declined() {
        let limits = ParseLimits {
            max_xml_events: 12_000,
            ..ParseLimits::default()
        };

        let package = parse_docx_s9_wire_with_limits(
            &many_chart_docx(4, 1_000),
            S9ParseOptions::default(),
            &limits,
        )
        .unwrap()
        .document
        .package;

        assert_eq!(package.document.content.len(), 4);
        assert_eq!(
            body_charts(&package.document.content),
            ["word/charts/chart1.xml", "word/charts/chart2.xml"]
        );
        assert_eq!(opaque_drawings(&package.document.content), 2);
    }

    #[test]
    fn a_damaged_chart_part_reads_as_an_absent_one() {
        let package = parse_docx_s9_wire(
            &chart_docx_with(CHART_PART, MALFORMED),
            S9ParseOptions::default(),
        )
        .unwrap()
        .document
        .package;
        let mut absent = chart_docx();
        absent.retain(|(path, _)| path != CHART_PART);
        let without = parse_docx_s9_wire(
            &ooxml_opc::rezip_parts(&absent).unwrap(),
            S9ParseOptions::default(),
        )
        .unwrap()
        .document
        .package;
        assert_eq!(package.document.content, without.document.content);
        assert_eq!(package.chart_entries, without.chart_entries);
        println!("DAMAGED == ABSENT");
    }
}
