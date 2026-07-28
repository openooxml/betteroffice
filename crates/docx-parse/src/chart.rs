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
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

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

pub fn parse_chart_parts(
    all_xml: &IndexMap<String, Vec<u8>>,
    budget: &mut ParseBudget<'_>,
) -> Result<ChartPartsMap, ParseError> {
    let mut charts = ChartPartsMap::new();
    for (path, xml) in all_xml {
        let normalized = normalize_chart_path(path)?;
        let lower = normalized.to_ascii_lowercase();
        if !lower.starts_with("word/charts/") || !lower.ends_with(".xml") {
            continue;
        }
        if let Some(chart) = parse_chart_xml(xml, Some(&normalized), path, budget)? {
            charts.insert(normalized.clone(), chart.clone());
            charts.insert(
                normalized
                    .strip_prefix("word/")
                    .unwrap_or(&normalized)
                    .to_owned(),
                chart,
            );
        }
    }
    Ok(charts)
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

pub fn parse_chart_from_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    charts: Option<&ChartPartsMap>,
) -> Result<Option<Chart>, ParseError> {
    let (Some(relationships), Some(charts)) = (relationships, charts) else {
        return Ok(None);
    };
    if charts.is_empty() {
        return Ok(None);
    }
    let Some(chart_ref) = first_deep(drawing, "chart", 0) else {
        return Ok(None);
    };
    let Some(relationship_id) = chart_ref.attribute(Some("r"), "id") else {
        return Ok(None);
    };
    let Some(relationship) = relationships.get(relationship_id) else {
        return Ok(None);
    };
    if relationship.relationship_type != relationship_types::CHART
        || relationship.target_mode == Some(TargetMode::External)
    {
        return Ok(None);
    }
    let path = normalize_chart_path(&relationship.target)?;
    let alias = path.strip_prefix("word/").unwrap_or(&path);
    let Some(source) = charts.get(&path).or_else(|| charts.get(alias)) else {
        return Ok(None);
    };
    let mut chart = source.clone();
    apply_drawing_metadata(&mut chart, drawing);
    chart.relationship_id = Some(relationship_id.to_owned());
    chart.path = Some(path);
    chart.size = parse_drawing_extent(drawing).or(chart.size);
    Ok(Some(chart))
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
    use crate::relationships::{Relationship, TargetMode};
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
        assert!(
            parse_chart_from_drawing(
                document.root().unwrap(),
                Some(&relationships),
                Some(&charts)
            )
            .unwrap()
            .is_none()
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
}
