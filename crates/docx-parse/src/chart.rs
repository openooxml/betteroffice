//! DrawingML chart parsing for the normalized display model.

use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::drawingml::{parse_color_element, resolve_color_value_to_hex};
use crate::image::{
    ImagePosition, ImageSize, ImageWrap, PositionAxis, parse_position_h, parse_position_v,
    parse_wrap_element_without_distances,
};
use crate::relationships::{
    RelationshipMap, TargetMode, relationship_types, resolve_relative_path,
};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

const DEFAULT_SERIES_COLORS: [&str; 8] = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E",
];
const MAX_DEEP_DEPTH: usize = 64;
const MAX_POINTS: usize = 100_000;
const MAX_PLOT_GROUPS: usize = 64;
const MAX_AXES: usize = 128;

pub type ChartPartsMap = IndexMap<String, Chart>;

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

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ChartLegend {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    pub visible: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub categories: Vec<String>,
    pub values: Vec<f64>,
    pub color: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub order: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category_formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value_formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_ids: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub points: Option<Vec<ChartPoint>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grouping: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker: Option<ChartMarker>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ChartMarker {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub symbol: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ChartPoint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub explosion: Option<f64>,
    pub color: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartPlotGroup {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grouping: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub overlap: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<f64>,
    pub axis_ids: Vec<String>,
    pub series: Vec<ChartSeries>,
    pub vary_colors: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_slice_angle: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<f64>,
    pub show_data_labels: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ChartAxes {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<ChartAxis>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<ChartAxis>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxis {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub labels: Option<Vec<String>>,
    pub axis_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cross_axis_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crosses: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub logarithmic_base: Option<f64>,
    pub reversed: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tick_label_position: Option<String>,
    pub hidden: bool,
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
    let Some(plot_area) = first_deep(chart_space, "plotArea", 0) else {
        return Ok(None);
    };
    let chart_elements = plot_area
        .child_elements()
        .filter(|child| plot_type_for(child).is_some())
        .take(MAX_PLOT_GROUPS)
        .collect::<Vec<_>>();
    if chart_elements.is_empty() {
        return Ok(None);
    }
    let plot_groups = chart_elements
        .into_iter()
        .map(parse_plot_group)
        .collect::<Vec<_>>();
    let first_type = plot_groups[0].chart_type.as_deref();
    let chart_type = match first_type {
        Some("bar" | "column" | "line" | "pie" | "doughnut") => first_type.unwrap().to_owned(),
        Some("ofPie") => "pie".to_owned(),
        _ => "line".to_owned(),
    };
    // Pinned normalization: legacy series deliberately discard every detailed
    // field except name/categories/values/color.
    let series = plot_groups
        .iter()
        .flat_map(|group| group.series.iter())
        .map(|series| ChartSeries {
            name: series.name.clone(),
            categories: series.categories.clone(),
            values: series.values.clone(),
            color: series.color.clone(),
            index: None,
            order: None,
            category_formula: None,
            value_formula: None,
            axis_ids: None,
            points: None,
            grouping: None,
            marker: None,
            smooth: None,
        })
        .collect::<Vec<_>>();
    let axis_list = plot_area
        .child_elements()
        .filter(|child| matches!(child.local_name(), "catAx" | "dateAx" | "valAx" | "serAx"))
        .take(MAX_AXES)
        .map(parse_axis)
        .collect::<Vec<_>>();
    let axes = parse_axes(plot_area, series.first());
    Ok(Some(Chart {
        content_type: "chart".to_owned(),
        chart_type,
        relationship_id: None,
        path: path.map(str::to_owned),
        title: chart_title(chart_space),
        legend: parse_legend(chart_space),
        series,
        axes,
        size: None,
        wrap: None,
        position: None,
        plot_groups,
        axis_list: (!axis_list.is_empty()).then_some(axis_list),
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

fn all_deep<'a>(root: &'a XmlElement, local: &str, depth: usize, output: &mut Vec<&'a XmlElement>) {
    if depth > MAX_DEEP_DEPTH || output.len() >= MAX_POINTS {
        return;
    }
    if root.local_name() == local {
        output.push(root);
    }
    for child in root.child_elements() {
        all_deep(child, local, depth + 1, output);
        if output.len() >= MAX_POINTS {
            break;
        }
    }
}

fn val_attr(element: Option<&XmlElement>) -> Option<&str> {
    let element = element?;
    element
        .attribute(None, "val")
        .or_else(|| element.attribute(Some("c"), "val"))
}

fn text_from_rich_text(parent: Option<&XmlElement>) -> Option<String> {
    let parent = parent?;
    if let Some(rich) = first_deep(parent, "rich", 0) {
        let mut elements = Vec::new();
        all_deep(rich, "t", 0, &mut elements);
        let text = elements
            .into_iter()
            .map(XmlElement::text_content)
            .collect::<String>();
        let trimmed = text.trim();
        if !trimmed.is_empty() {
            return Some(trimmed.to_owned());
        }
    }
    let text = first_deep(parent, "v", 0)
        .map(XmlElement::text_content)
        .unwrap_or_default();
    nonempty_trimmed(&text)
}

fn chart_title(chart_space: &XmlElement) -> Option<String> {
    text_from_rich_text(first_deep(chart_space, "title", 0))
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

fn parse_string_cache(parent: Option<&XmlElement>) -> Vec<String> {
    let Some(parent) = parent else {
        return Vec::new();
    };
    let Some(cache) = first_deep(parent, "strCache", 0)
        .or_else(|| first_deep(parent, "multiLvlStrCache", 0))
        .or_else(|| first_deep(parent, "numCache", 0))
    else {
        return Vec::new();
    };
    cache
        .children_by_local_name("pt")
        .take(MAX_POINTS)
        .map(|point| {
            point
                .child_by_local_name("v")
                .map(XmlElement::text_content)
                .unwrap_or_default()
                .trim()
                .to_owned()
        })
        .collect()
}

fn parse_num_cache(parent: Option<&XmlElement>) -> Vec<f64> {
    let Some(parent) = parent else {
        return Vec::new();
    };
    let Some(cache) = first_deep(parent, "numCache", 0) else {
        return Vec::new();
    };
    cache
        .children_by_local_name("pt")
        .take(MAX_POINTS)
        .filter_map(|point| {
            let text = point.child_by_local_name("v")?.text_content();
            parse_number(Some(text.trim()))
        })
        .collect()
}

fn parse_series_name(series: &XmlElement) -> Option<String> {
    text_from_rich_text(series.child_by_local_name("tx"))
}

fn parse_series_color(series: &XmlElement, index: usize) -> String {
    let parsed = series
        .child_by_local_name("spPr")
        .and_then(|properties| first_deep(properties, "solidFill", 0))
        .and_then(|fill| resolve_color_value_to_hex(parse_color_element(Some(fill)).as_ref()));
    parsed.unwrap_or_else(|| DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.len()].to_owned())
}

fn parse_series(chart: &XmlElement, grouping: Option<&str>) -> Vec<ChartSeries> {
    chart
        .children_by_local_name("ser")
        .enumerate()
        .map(|(index, series)| {
            let category = series.child_by_local_name("cat");
            let value = series.child_by_local_name("val");
            let marker = series.child_by_local_name("marker");
            let marker_symbol =
                val_attr(marker.and_then(|value| value.child_by_local_name("symbol")))
                    .map(str::to_owned);
            let marker_size = parse_number(val_attr(
                marker.and_then(|value| value.child_by_local_name("size")),
            ));
            let axis_ids = chart
                .children_by_local_name("axId")
                .filter_map(|axis| {
                    val_attr(Some(axis))
                        .filter(|value| !value.is_empty())
                        .map(str::to_owned)
                })
                .collect::<Vec<_>>();
            let points = series
                .children_by_local_name("dPt")
                .take(MAX_POINTS)
                .map(|point| ChartPoint {
                    index: parse_number(val_attr(point.child_by_local_name("idx"))),
                    explosion: parse_number(val_attr(point.child_by_local_name("explosion"))),
                    color: parse_series_color(point, index),
                })
                .collect::<Vec<_>>();
            ChartSeries {
                name: parse_series_name(series),
                categories: parse_string_cache(category),
                values: parse_num_cache(value),
                color: parse_series_color(series, index),
                index: parse_number(val_attr(series.child_by_local_name("idx"))),
                order: parse_number(val_attr(series.child_by_local_name("order"))),
                category_formula: child_formula(category),
                value_formula: child_formula(value),
                axis_ids: (!axis_ids.is_empty()).then_some(axis_ids),
                points: (!points.is_empty()).then_some(points),
                grouping: grouping.map(str::to_owned),
                marker: (marker_symbol.is_some() || marker_size.is_some()).then_some(ChartMarker {
                    symbol: marker_symbol,
                    size: marker_size,
                }),
                smooth: (val_attr(series.child_by_local_name("smooth")) == Some("1"))
                    .then_some(true),
            }
        })
        .collect()
}

fn child_formula(parent: Option<&XmlElement>) -> Option<String> {
    let text = first_deep(parent?, "f", 0)?.text_content();
    nonempty_trimmed(&text)
}

fn parse_legend(chart_space: &XmlElement) -> Option<ChartLegend> {
    let legend = first_deep(chart_space, "legend", 0)?;
    let position = match val_attr(legend.child_by_local_name("legendPos")) {
        Some("l") => Some("left"),
        Some("r") => Some("right"),
        Some("t") => Some("top"),
        Some("b") => Some("bottom"),
        _ => None,
    };
    Some(ChartLegend {
        position: position.map(str::to_owned),
        visible: true,
    })
}

fn parse_axis(axis: &XmlElement) -> ChartAxis {
    let scaling = axis.child_by_local_name("scaling");
    let crosses = val_attr(axis.child_by_local_name("crosses"));
    ChartAxis {
        id: val_attr(axis.child_by_local_name("axId")).map(str::to_owned),
        title: text_from_rich_text(axis.child_by_local_name("title")),
        min: parse_number(val_attr(
            scaling.and_then(|value| value.child_by_local_name("min")),
        )),
        max: parse_number(val_attr(
            scaling.and_then(|value| value.child_by_local_name("max")),
        )),
        labels: None,
        axis_type: match axis.local_name() {
            "catAx" => "category",
            "dateAx" => "date",
            "serAx" => "series",
            _ => "value",
        }
        .to_owned(),
        position: match val_attr(axis.child_by_local_name("axPos")) {
            Some("l") => Some("left"),
            Some("r") => Some("right"),
            Some("t") => Some("top"),
            Some("b") => Some("bottom"),
            _ => None,
        }
        .map(str::to_owned),
        cross_axis_id: val_attr(axis.child_by_local_name("crossAx")).map(str::to_owned),
        crosses: crosses
            .filter(|value| matches!(*value, "min" | "max" | "autoZero"))
            .map(str::to_owned),
        crosses_at: parse_number(val_attr(axis.child_by_local_name("crossesAt"))),
        major_unit: parse_number(val_attr(axis.child_by_local_name("majorUnit"))),
        minor_unit: parse_number(val_attr(axis.child_by_local_name("minorUnit"))),
        logarithmic_base: parse_number(val_attr(
            scaling.and_then(|value| value.child_by_local_name("logBase")),
        )),
        reversed: val_attr(scaling.and_then(|value| value.child_by_local_name("orientation")))
            == Some("maxMin"),
        number_format: axis
            .child_by_local_name("numFmt")
            .and_then(|value| value.attribute(None, "formatCode"))
            .map(str::to_owned),
        major_tick_mark: val_attr(axis.child_by_local_name("majorTickMark")).map(str::to_owned),
        minor_tick_mark: val_attr(axis.child_by_local_name("minorTickMark")).map(str::to_owned),
        tick_label_position: val_attr(axis.child_by_local_name("tickLblPos")).map(str::to_owned),
        hidden: val_attr(axis.child_by_local_name("delete")) == Some("1"),
    }
}

fn parse_axes(plot_area: &XmlElement, first_series: Option<&ChartSeries>) -> Option<ChartAxes> {
    let category = plot_area
        .child_by_local_name("catAx")
        .or_else(|| plot_area.child_by_local_name("dateAx"))
        .map(parse_axis);
    let value = plot_area.child_by_local_name("valAx").map(parse_axis);
    let mut category = category;
    if let (Some(axis), Some(series)) = (&mut category, first_series)
        && !series.categories.is_empty()
    {
        axis.labels = Some(series.categories.clone());
    }
    (category.is_some() || value.is_some()).then_some(ChartAxes { category, value })
}

fn plot_type_for(chart: &XmlElement) -> Option<String> {
    let local = chart.local_name().replace("3DChart", "Chart");
    let value = match local.as_str() {
        "barChart" => {
            if val_attr(chart.child_by_local_name("barDir")) == Some("bar") {
                "bar"
            } else {
                "column"
            }
        }
        "lineChart" => "line",
        "pieChart" => "pie",
        "doughnutChart" => "doughnut",
        "areaChart" => "area",
        "scatterChart" => "scatter",
        "radarChart" => "radar",
        "stockChart" => "stock",
        "bubbleChart" => "bubble",
        "ofPieChart" => "ofPie",
        "surfaceChart" => "surface",
        _ => return None,
    };
    Some(value.to_owned())
}

fn parse_grouping(chart: &XmlElement) -> Option<String> {
    match val_attr(chart.child_by_local_name("grouping")) {
        Some(value @ ("stacked" | "percentStacked" | "clustered" | "standard")) => {
            Some(value.to_owned())
        }
        _ => None,
    }
}

fn parse_plot_group(chart: &XmlElement) -> ChartPlotGroup {
    let grouping = parse_grouping(chart);
    ChartPlotGroup {
        chart_type: plot_type_for(chart),
        grouping: grouping.clone(),
        overlap: parse_number(val_attr(chart.child_by_local_name("overlap"))),
        gap_width: parse_number(val_attr(chart.child_by_local_name("gapWidth"))),
        axis_ids: chart
            .children_by_local_name("axId")
            .filter_map(|axis| {
                val_attr(Some(axis))
                    .filter(|value| !value.is_empty())
                    .map(str::to_owned)
            })
            .collect(),
        series: parse_series(chart, grouping.as_deref()),
        vary_colors: val_attr(chart.child_by_local_name("varyColors")) == Some("1"),
        first_slice_angle: parse_number(val_attr(chart.child_by_local_name("firstSliceAng"))),
        hole_size: parse_number(val_attr(chart.child_by_local_name("holeSize"))),
        show_data_labels: chart.child_by_local_name("dLbls").is_some(),
    }
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

fn nonempty_trimmed(value: &str) -> Option<String> {
    let value = value.trim();
    (!value.is_empty()).then(|| value.to_owned())
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
