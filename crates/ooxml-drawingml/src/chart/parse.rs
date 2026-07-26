//! `c:chartSpace` parsing, generic over the host's XML element type.

use super::model::{
    ChartAxes, ChartAxis, ChartLegend, ChartMarker, ChartPlotGroup, ChartPoint, ChartSeries,
    ChartSpace,
};

pub const DEFAULT_SERIES_COLORS: [&str; 8] = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E",
];
const MAX_DEEP_DEPTH: usize = 64;
const MAX_POINTS: usize = 100_000;
const MAX_PLOT_GROUPS: usize = 64;
const MAX_AXES: usize = 128;
/// Chart-wide, so per-vector limits cannot multiply into an unbounded parse.
const MAX_CHART_SERIES: usize = 1_024;
const MAX_CHART_POINTS: usize = 200_000;
const MAX_AXIS_IDS: usize = 16;

/// What one `c:chartSpace` may still allocate, shared across its plot groups.
struct Budget {
    series: usize,
    points: usize,
}

impl Budget {
    fn new() -> Self {
        Self {
            series: MAX_CHART_SERIES,
            points: MAX_CHART_POINTS,
        }
    }

    fn series_cap(&self) -> usize {
        self.series
    }

    fn spend_series(&mut self, used: usize) {
        self.series = self.series.saturating_sub(used);
    }

    fn point_cap(&self, requested: usize) -> usize {
        requested.min(self.points)
    }

    fn spend_points(&mut self, used: usize) {
        self.points = self.points.saturating_sub(used);
    }
}

/// Read-only XML access [`parse_chart_space`] needs. Hosts implement it for
/// their own element type, or for an adapter carrying host context such as a
/// theme, so the chart parser stays format-agnostic.
pub trait ChartXml: Sized {
    /// The local name, namespace prefix stripped: `ser` for a `<c:ser>`.
    fn local_name(&self) -> &str;
    /// The `prefix:name` attribute when `prefix` is given, falling back to the
    /// unqualified `name`; `None` looks up `name` alone.
    fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str>;
    /// Direct child elements in document order. Text nodes never appear.
    fn child_elements(&self) -> impl Iterator<Item = &Self>;
    /// Every descendant text and CDATA node, concatenated in document order.
    fn descendant_text(&self) -> String;
    /// `#RRGGBB` for an `a:solidFill` element, resolved through the host's own
    /// theme and color modifiers.
    fn solid_fill_hex(&self) -> Option<String>;
}

/// Parse a `c:chartSpace` root. `None` when it carries no recognized plot.
pub fn parse_chart_space<E: ChartXml>(chart_space: &E) -> Option<ChartSpace> {
    let plot_area = first_deep(chart_space, "plotArea", 0)?;
    let chart_elements = plot_area
        .child_elements()
        .filter(|child| plot_type_for(*child).is_some())
        .take(MAX_PLOT_GROUPS)
        .collect::<Vec<_>>();
    if chart_elements.is_empty() {
        return None;
    }
    let budget = &mut Budget::new();
    let plot_groups = chart_elements
        .into_iter()
        .map(|chart| parse_plot_group(chart, budget))
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
    Some(ChartSpace {
        chart_type,
        title: chart_title(chart_space),
        legend: parse_legend(chart_space),
        series,
        axes,
        plot_groups,
        axis_list: (!axis_list.is_empty()).then_some(axis_list),
    })
}

fn child<'a, E: ChartXml>(parent: &'a E, local: &str) -> Option<&'a E> {
    parent
        .child_elements()
        .find(|node| node.local_name() == local)
}

fn children<'a, E: ChartXml>(parent: &'a E, local: &'a str) -> impl Iterator<Item = &'a E> {
    parent
        .child_elements()
        .filter(move |node| node.local_name() == local)
}

fn first_deep<'a, E: ChartXml>(root: &'a E, local: &str, depth: usize) -> Option<&'a E> {
    if depth > MAX_DEEP_DEPTH {
        return None;
    }
    if root.local_name() == local {
        return Some(root);
    }
    root.child_elements()
        .find_map(|node| first_deep(node, local, depth + 1))
}

fn all_deep<'a, E: ChartXml>(root: &'a E, local: &str, depth: usize, output: &mut Vec<&'a E>) {
    if depth > MAX_DEEP_DEPTH || output.len() >= MAX_POINTS {
        return;
    }
    if root.local_name() == local {
        output.push(root);
    }
    for node in root.child_elements() {
        all_deep(node, local, depth + 1, output);
        if output.len() >= MAX_POINTS {
            break;
        }
    }
}

fn val_attr<E: ChartXml>(element: Option<&E>) -> Option<&str> {
    let element = element?;
    element
        .attribute(None, "val")
        .or_else(|| element.attribute(Some("c"), "val"))
}

fn text_from_rich_text<E: ChartXml>(parent: Option<&E>) -> Option<String> {
    let parent = parent?;
    if let Some(rich) = first_deep(parent, "rich", 0) {
        let mut elements = Vec::new();
        all_deep(rich, "t", 0, &mut elements);
        let text = elements
            .into_iter()
            .map(E::descendant_text)
            .collect::<String>();
        let trimmed = text.trim();
        if !trimmed.is_empty() {
            return Some(trimmed.to_owned());
        }
    }
    let text = first_deep(parent, "v", 0)
        .map(E::descendant_text)
        .unwrap_or_default();
    nonempty_trimmed(&text)
}

fn chart_title<E: ChartXml>(chart_space: &E) -> Option<String> {
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

/// A schema `xsd:unsignedInt` index: decimal digits in `u32` range. Signed,
/// fractional, exponent and radix-prefixed forms are not indexes, so they never
/// reach a host's `as usize`.
fn parse_index(raw: Option<&str>) -> Option<f64> {
    let value = raw?.trim();
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    value.parse::<u32>().ok().map(f64::from)
}

/// Reads at most the remaining point budget from `elements`, charging every
/// child it examines so malformed ones cost as much as parsed ones.
fn take_points<'a, E: ChartXml + 'a, T>(
    elements: impl Iterator<Item = &'a E>,
    budget: &mut Budget,
    mut parse: impl FnMut(&'a E) -> Option<T>,
) -> Vec<T> {
    let mut examined = 0;
    let values = elements
        .take(budget.point_cap(MAX_POINTS))
        .filter_map(|element| {
            examined += 1;
            parse(element)
        })
        .collect::<Vec<_>>();
    budget.spend_points(examined);
    values
}

fn parse_string_cache<E: ChartXml>(parent: Option<&E>, budget: &mut Budget) -> Vec<String> {
    let Some(parent) = parent else {
        return Vec::new();
    };
    let Some(cache) = first_deep(parent, "strCache", 0)
        .or_else(|| first_deep(parent, "multiLvlStrCache", 0))
        .or_else(|| first_deep(parent, "numCache", 0))
    else {
        return Vec::new();
    };
    take_points(children(cache, "pt"), budget, |point| {
        Some(
            child(point, "v")
                .map(E::descendant_text)
                .unwrap_or_default()
                .trim()
                .to_owned(),
        )
    })
}

fn parse_num_cache<E: ChartXml>(parent: Option<&E>, budget: &mut Budget) -> Vec<f64> {
    let Some(parent) = parent else {
        return Vec::new();
    };
    let Some(cache) = first_deep(parent, "numCache", 0) else {
        return Vec::new();
    };
    take_points(children(cache, "pt"), budget, |point| {
        let text = child(point, "v")?.descendant_text();
        parse_number(Some(text.trim()))
    })
}

fn parse_series_name<E: ChartXml>(series: &E) -> Option<String> {
    text_from_rich_text(child(series, "tx"))
}

fn parse_series_color<E: ChartXml>(series: &E, index: usize) -> String {
    let parsed = child(series, "spPr")
        .and_then(|properties| first_deep(properties, "solidFill", 0))
        .and_then(E::solid_fill_hex);
    parsed.unwrap_or_else(|| DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.len()].to_owned())
}

fn parse_series<E: ChartXml>(
    chart: &E,
    grouping: Option<&str>,
    axis_ids: &[String],
    budget: &mut Budget,
) -> Vec<ChartSeries> {
    let cap = budget.series_cap();
    let series = children(chart, "ser")
        .enumerate()
        .take(cap)
        .map(|(index, series)| {
            let category = child(series, "cat");
            let value = child(series, "val");
            let marker = child(series, "marker");
            let marker_symbol =
                val_attr(marker.and_then(|value| child(value, "symbol"))).map(str::to_owned);
            let marker_size = parse_number(val_attr(marker.and_then(|value| child(value, "size"))));
            let points = take_points(children(series, "dPt"), budget, |point| {
                let point_index = match child(point, "idx") {
                    Some(idx) => Some(parse_index(val_attr(Some(idx)))?),
                    None => None,
                };
                Some(ChartPoint {
                    index: point_index,
                    explosion: parse_number(val_attr(child(point, "explosion"))),
                    color: parse_series_color(point, index),
                })
            });
            ChartSeries {
                name: parse_series_name(series),
                categories: parse_string_cache(category, budget),
                values: parse_num_cache(value, budget),
                color: parse_series_color(series, index),
                index: parse_index(val_attr(child(series, "idx"))),
                order: parse_index(val_attr(child(series, "order"))),
                category_formula: child_formula(category),
                value_formula: child_formula(value),
                axis_ids: (!axis_ids.is_empty()).then(|| axis_ids.to_vec()),
                points: (!points.is_empty()).then_some(points),
                grouping: grouping.map(str::to_owned),
                marker: (marker_symbol.is_some() || marker_size.is_some()).then_some(ChartMarker {
                    symbol: marker_symbol,
                    size: marker_size,
                }),
                smooth: (val_attr(child(series, "smooth")) == Some("1")).then_some(true),
            }
        })
        .collect::<Vec<_>>();
    budget.spend_series(series.len());
    series
}

fn child_formula<E: ChartXml>(parent: Option<&E>) -> Option<String> {
    let text = first_deep(parent?, "f", 0)?.descendant_text();
    nonempty_trimmed(&text)
}

fn parse_legend<E: ChartXml>(chart_space: &E) -> Option<ChartLegend> {
    let legend = first_deep(chart_space, "legend", 0)?;
    let position = match val_attr(child(legend, "legendPos")) {
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

fn parse_axis<E: ChartXml>(axis: &E) -> ChartAxis {
    let scaling = child(axis, "scaling");
    let crosses = val_attr(child(axis, "crosses"));
    ChartAxis {
        id: val_attr(child(axis, "axId")).map(str::to_owned),
        title: text_from_rich_text(child(axis, "title")),
        min: parse_number(val_attr(scaling.and_then(|value| child(value, "min")))),
        max: parse_number(val_attr(scaling.and_then(|value| child(value, "max")))),
        labels: None,
        axis_type: match axis.local_name() {
            "catAx" => "category",
            "dateAx" => "date",
            "serAx" => "series",
            _ => "value",
        }
        .to_owned(),
        position: match val_attr(child(axis, "axPos")) {
            Some("l") => Some("left"),
            Some("r") => Some("right"),
            Some("t") => Some("top"),
            Some("b") => Some("bottom"),
            _ => None,
        }
        .map(str::to_owned),
        cross_axis_id: val_attr(child(axis, "crossAx")).map(str::to_owned),
        crosses: crosses
            .filter(|value| matches!(*value, "min" | "max" | "autoZero"))
            .map(str::to_owned),
        crosses_at: parse_number(val_attr(child(axis, "crossesAt"))),
        major_unit: parse_number(val_attr(child(axis, "majorUnit"))),
        minor_unit: parse_number(val_attr(child(axis, "minorUnit"))),
        logarithmic_base: parse_number(val_attr(scaling.and_then(|value| child(value, "logBase")))),
        reversed: val_attr(scaling.and_then(|value| child(value, "orientation"))) == Some("maxMin"),
        number_format: child(axis, "numFmt")
            .and_then(|value| value.attribute(None, "formatCode"))
            .map(str::to_owned),
        major_tick_mark: val_attr(child(axis, "majorTickMark")).map(str::to_owned),
        minor_tick_mark: val_attr(child(axis, "minorTickMark")).map(str::to_owned),
        tick_label_position: val_attr(child(axis, "tickLblPos")).map(str::to_owned),
        hidden: val_attr(child(axis, "delete")) == Some("1"),
    }
}

fn parse_axes<E: ChartXml>(plot_area: &E, first_series: Option<&ChartSeries>) -> Option<ChartAxes> {
    let category = child(plot_area, "catAx")
        .or_else(|| child(plot_area, "dateAx"))
        .map(parse_axis);
    let value = child(plot_area, "valAx").map(parse_axis);
    let mut category = category;
    if let (Some(axis), Some(series)) = (&mut category, first_series)
        && !series.categories.is_empty()
    {
        axis.labels = Some(series.categories.clone());
    }
    (category.is_some() || value.is_some()).then_some(ChartAxes { category, value })
}

fn plot_type_for<E: ChartXml>(chart: &E) -> Option<String> {
    let local = chart.local_name().replace("3DChart", "Chart");
    let value = match local.as_str() {
        "barChart" => {
            if val_attr(child(chart, "barDir")) == Some("bar") {
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

fn parse_grouping<E: ChartXml>(chart: &E) -> Option<String> {
    match val_attr(child(chart, "grouping")) {
        Some(value @ ("stacked" | "percentStacked" | "clustered" | "standard")) => {
            Some(value.to_owned())
        }
        _ => None,
    }
}

fn parse_plot_group<E: ChartXml>(chart: &E, budget: &mut Budget) -> ChartPlotGroup {
    let grouping = parse_grouping(chart);
    let axis_ids = children(chart, "axId")
        .filter_map(|axis| {
            val_attr(Some(axis))
                .filter(|value| !value.is_empty())
                .map(str::to_owned)
        })
        .take(MAX_AXIS_IDS)
        .collect::<Vec<_>>();
    ChartPlotGroup {
        chart_type: plot_type_for(chart),
        grouping: grouping.clone(),
        overlap: parse_number(val_attr(child(chart, "overlap"))),
        gap_width: parse_number(val_attr(child(chart, "gapWidth"))),
        series: parse_series(chart, grouping.as_deref(), &axis_ids, budget),
        axis_ids,
        vary_colors: val_attr(child(chart, "varyColors")) == Some("1"),
        first_slice_angle: parse_number(val_attr(child(chart, "firstSliceAng"))),
        hole_size: parse_number(val_attr(child(chart, "holeSize"))),
        show_data_labels: child(chart, "dLbls").is_some(),
    }
}

fn nonempty_trimmed(value: &str) -> Option<String> {
    let value = value.trim();
    (!value.is_empty()).then(|| value.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::{PlotChart, PlotOp, PlotRect, plot_chart};

    #[derive(Clone)]
    struct Node {
        name: String,
        attrs: Vec<(String, String)>,
        text: String,
        children: Vec<Node>,
    }

    impl Node {
        fn el(name: &str, children: Vec<Node>) -> Self {
            Self {
                name: name.to_owned(),
                attrs: Vec::new(),
                text: String::new(),
                children,
            }
        }

        fn val(name: &str, value: &str) -> Self {
            Self::el(name, Vec::new()).attr("val", value)
        }

        fn text(name: &str, value: &str) -> Self {
            Self {
                name: name.to_owned(),
                attrs: Vec::new(),
                text: value.to_owned(),
                children: Vec::new(),
            }
        }

        fn attr(mut self, key: &str, value: &str) -> Self {
            self.attrs.push((key.to_owned(), value.to_owned()));
            self
        }
    }

    impl ChartXml for Node {
        fn local_name(&self) -> &str {
            self.name
                .split_once(':')
                .map_or(self.name.as_str(), |(_, local)| local)
        }

        fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
            prefix
                .and_then(|prefix| {
                    let qualified = format!("{prefix}:{name}");
                    self.attrs.iter().find(|(key, _)| *key == qualified)
                })
                .or_else(|| self.attrs.iter().find(|(key, _)| key.as_str() == name))
                .map(|(_, value)| value.as_str())
        }

        fn child_elements(&self) -> impl Iterator<Item = &Self> {
            self.children.iter()
        }

        fn descendant_text(&self) -> String {
            let mut text = self.text.clone();
            for child in &self.children {
                text.push_str(&child.descendant_text());
            }
            text
        }

        fn solid_fill_hex(&self) -> Option<String> {
            self.children
                .first()
                .and_then(|child| child.attribute(None, "val"))
                .map(|rgb| format!("#{rgb}"))
        }
    }

    /// The same chart, with every element name qualified by `ns`.
    fn fixture(ns: &str) -> Node {
        let el = |name: &str, children: Vec<Node>| Node::el(&format!("{ns}{name}"), children);
        let val = |name: &str, value: &str| Node::val(&format!("{ns}{name}"), value);
        let text = |name: &str, value: &str| Node::text(&format!("{ns}{name}"), value);
        el(
            "chartSpace",
            vec![el(
                "chart",
                vec![
                    el("title", vec![el("rich", vec![text("t", " Sales ")])]),
                    el("legend", vec![val("legendPos", "l")]),
                    el(
                        "plotArea",
                        vec![
                            el(
                                "barChart",
                                vec![
                                    val("barDir", "bar"),
                                    val("grouping", "stacked"),
                                    el(
                                        "ser",
                                        vec![
                                            el(
                                                "tx",
                                                vec![el(
                                                    "v",
                                                    vec![text("x", "Nor"), text("x", "th")],
                                                )],
                                            ),
                                            el(
                                                "spPr",
                                                vec![el(
                                                    "solidFill",
                                                    vec![val("srgbClr", "FF0000")],
                                                )],
                                            ),
                                            el(
                                                "cat",
                                                vec![el(
                                                    "strCache",
                                                    vec![el("pt", vec![text("v", "Q1")])],
                                                )],
                                            ),
                                            el(
                                                "val",
                                                vec![el(
                                                    "numCache",
                                                    vec![el("pt", vec![text("v", "7")])],
                                                )],
                                            ),
                                        ],
                                    ),
                                    val("axId", "1"),
                                ],
                            ),
                            el("catAx", vec![val("axId", "1"), val("axPos", "b")]),
                            el(
                                "valAx",
                                vec![
                                    val("axId", "2"),
                                    el("scaling", vec![val("min", "0"), val("max", "9")]),
                                ],
                            ),
                        ],
                    ),
                ],
            )],
        )
    }

    #[test]
    fn parses_series_axes_and_legend_from_a_generic_element_tree() {
        let parsed = parse_chart_space(&fixture("")).expect("chart space parses");
        assert_eq!(parsed.chart_type, "bar");
        assert_eq!(parsed.title.as_deref(), Some("Sales"));
        assert_eq!(parsed.legend.unwrap().position.as_deref(), Some("left"));
        assert_eq!(parsed.series[0].name.as_deref(), Some("North"));
        assert_eq!(parsed.series[0].categories, ["Q1"]);
        assert_eq!(parsed.series[0].values, [7.0]);
        assert_eq!(parsed.series[0].color, "#FF0000");
        assert_eq!(parsed.plot_groups[0].grouping.as_deref(), Some("stacked"));
        assert_eq!(parsed.axis_list.as_ref().unwrap().len(), 2);
        let axes = parsed.axes.unwrap();
        assert_eq!(axes.category.unwrap().labels.unwrap(), ["Q1"]);
        assert_eq!(axes.value.unwrap().max, Some(9.0));
    }

    #[test]
    fn a_namespace_prefixed_document_parses_the_same_as_a_bare_one() {
        let bare = parse_chart_space(&fixture("")).expect("bare parses");
        let prefixed = parse_chart_space(&fixture("c:")).expect("prefixed parses");
        assert_eq!(bare, prefixed);
        assert_eq!(prefixed.series[0].name.as_deref(), Some("North"));
    }

    #[test]
    fn a_qualified_val_attribute_resolves_through_the_prefix_fallback() {
        let space = Node::el(
            "c:chartSpace",
            vec![Node::el(
                "c:chart",
                vec![
                    Node::el(
                        "c:legend",
                        vec![Node::el("c:legendPos", Vec::new()).attr("c:val", "b")],
                    ),
                    Node::el("c:plotArea", vec![Node::el("c:pieChart", Vec::new())]),
                ],
            )],
        );
        let parsed = parse_chart_space(&space).expect("chart space parses");
        assert_eq!(parsed.legend.unwrap().position.as_deref(), Some("bottom"));
    }

    #[test]
    fn chart_wide_budgets_cap_series_and_axis_ids() {
        let mut bar = vec![Node::val("c:barDir", "col")];
        bar.extend((0..2_000).map(|_| {
            Node::el(
                "c:ser",
                vec![Node::el("c:tx", vec![Node::text("c:v", "S")])],
            )
        }));
        bar.extend((0..64).map(|_| Node::val("c:axId", "1")));
        let space = Node::el(
            "c:chartSpace",
            vec![Node::el(
                "c:chart",
                vec![Node::el("c:plotArea", vec![Node::el("c:barChart", bar)])],
            )],
        );

        let parsed = parse_chart_space(&space).expect("chart space parses");
        assert_eq!(parsed.plot_groups[0].series.len(), MAX_CHART_SERIES);
        assert_eq!(parsed.plot_groups[0].axis_ids.len(), MAX_AXIS_IDS);
        assert!(
            parsed.plot_groups[0].series.iter().all(|series| series
                .axis_ids
                .as_ref()
                .unwrap()
                .len()
                == MAX_AXIS_IDS)
        );
    }

    #[test]
    fn out_of_schema_indexes_are_not_parsed_as_indexes() {
        assert_eq!(parse_index(Some("2")), Some(2.0));
        assert_eq!(parse_index(Some("0")), Some(0.0));
        assert_eq!(parse_index(Some("007")), Some(7.0));
        assert_eq!(parse_index(Some("4294967295")), Some(4_294_967_295.0));
        assert_eq!(parse_index(Some("4294967296")), None);
        assert_eq!(parse_index(Some("-1")), None);
        assert_eq!(parse_index(Some("+1")), None);
        assert_eq!(parse_index(Some("1.5")), None);
        assert_eq!(parse_index(Some("2.0")), None);
        assert_eq!(parse_index(Some("1e3")), None);
        assert_eq!(parse_index(Some("1e30")), None);
        assert_eq!(parse_index(Some("0x2")), None);
        assert_eq!(parse_index(Some("nonsense")), None);
    }

    /// A `c:dPt` carrying `idx`, if given, and a red fill.
    fn red_point(index: Option<&str>) -> Node {
        let mut children = index
            .map(|index| vec![Node::val("c:idx", index)])
            .unwrap_or_default();
        children.push(Node::el(
            "c:spPr",
            vec![Node::el(
                "c:solidFill",
                vec![Node::val("a:srgbClr", "FF0000")],
            )],
        ));
        Node::el("c:dPt", children)
    }

    /// A two-slice pie chart carrying `points`.
    fn pie_with(points: Vec<Node>) -> Node {
        let mut series = vec![Node::el(
            "c:val",
            vec![Node::el(
                "c:numCache",
                vec![
                    Node::el("c:pt", vec![Node::text("c:v", "3")]),
                    Node::el("c:pt", vec![Node::text("c:v", "1")]),
                ],
            )],
        )];
        series.extend(points);
        Node::el(
            "c:chartSpace",
            vec![Node::el(
                "c:chart",
                vec![Node::el(
                    "c:plotArea",
                    vec![Node::el("c:pieChart", vec![Node::el("c:ser", series)])],
                )],
            )],
        )
    }

    fn red_wedges(space: &ChartSpace) -> usize {
        let rect = PlotRect {
            x: 0.0,
            y: 0.0,
            w: 300.0,
            h: 200.0,
        };
        plot_chart(&PlotChart::from(space), rect)
            .iter()
            .filter(|op| matches!(op, PlotOp::Path { fill, .. } if fill == "#FF0000"))
            .count()
    }

    #[test]
    fn an_invalid_point_index_is_dropped_instead_of_matching_every_slice() {
        for index in ["-1", "1.5", "1e3", "0x2", "", "nonsense"] {
            let space = parse_chart_space(&pie_with(vec![red_point(Some(index))]))
                .expect("chart space parses");
            assert!(space.plot_groups[0].series[0].points.is_none(), "{index}");
            assert_eq!(red_wedges(&space), 0, "{index}");
        }

        let space = parse_chart_space(&pie_with(vec![red_point(Some("-1")), red_point(Some("1"))]))
            .expect("chart space parses");
        let points = space.plot_groups[0].series[0].points.as_ref().unwrap();
        assert_eq!(points.len(), 1);
        assert_eq!(points[0].index, Some(1.0));
        assert_eq!(red_wedges(&space), 1);
    }

    #[test]
    fn an_absent_point_index_stays_the_intentional_wildcard() {
        let space =
            parse_chart_space(&pie_with(vec![red_point(None)])).expect("chart space parses");
        let points = space.plot_groups[0].series[0].points.as_ref().unwrap();
        assert_eq!(points[0].index, None);
        assert_eq!(red_wedges(&space), 2);
    }

    #[test]
    fn every_examined_point_costs_budget_even_when_it_yields_nothing() {
        let cache = Node::el(
            "c:numCache",
            vec![
                Node::el("c:pt", vec![Node::text("c:v", "nonsense")]),
                Node::el("c:pt", vec![Node::text("c:v", "1")]),
                Node::el("c:pt", Vec::new()),
            ],
        );
        let mut budget = Budget::new();
        assert_eq!(parse_num_cache(Some(&cache), &mut budget), [1.0]);
        assert_eq!(budget.point_cap(MAX_CHART_POINTS), MAX_CHART_POINTS - 3);

        let chart = Node::el(
            "c:pieChart",
            vec![Node::el(
                "c:ser",
                vec![
                    Node::el("c:dPt", vec![Node::val("c:idx", "-1")]),
                    Node::el("c:dPt", vec![Node::val("c:idx", "1")]),
                ],
            )],
        );
        let mut budget = Budget::new();
        let series = parse_series(&chart, None, &[], &mut budget);
        assert_eq!(series[0].points.as_ref().unwrap().len(), 1);
        assert_eq!(budget.point_cap(MAX_CHART_POINTS), MAX_CHART_POINTS - 2);
    }

    #[test]
    fn the_point_budget_is_shared_across_every_cache_in_one_chart() {
        let mut budget = Budget::new();
        assert_eq!(budget.point_cap(MAX_POINTS), MAX_POINTS);
        budget.spend_points(MAX_CHART_POINTS - 5);
        assert_eq!(budget.point_cap(MAX_POINTS), 5);
        budget.spend_points(9);
        assert_eq!(budget.point_cap(MAX_POINTS), 0);
    }

    #[test]
    fn returns_none_without_a_recognized_plot() {
        assert!(
            parse_chart_space(&Node::el(
                "c:chartSpace",
                vec![Node::el("c:chart", Vec::new())]
            ))
            .is_none()
        );
        assert!(
            parse_chart_space(&Node::el(
                "c:chartSpace",
                vec![Node::el(
                    "c:chart",
                    vec![Node::el("c:plotArea", Vec::new())]
                )]
            ))
            .is_none()
        );
    }
}
