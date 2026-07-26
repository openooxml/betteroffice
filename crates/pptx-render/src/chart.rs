//! Chart plotting: shared plot ops translated into slide primitives.

use std::collections::HashMap;

use ooxml_drawingml::GeometryPathCommand;
use ooxml_drawingml::chart::{
    ChartSpace, PlotChart, PlotFont, PlotOp, PlotPoint, PlotRect, PlotSink, chart_aria_label,
    format_number, plot_chart_into,
};
use pptx_parse::{ChartDataLabels, ChartPlotGroup, ChartPointLabel, ChartSeries};

use crate::{Paint, Primitive, RenderError, Stroke, Transform};

/// Value labels one chart may contribute, shared across its series, so a
/// chart with thousands of categories cannot turn every one into a text op.
const MAX_DATA_LABELS: usize = 4_096;

/// The graphic frame a chart is drawn into.
pub(crate) struct ChartFrame<'a> {
    pub object_id: u32,
    pub shape_id: Option<&'a str>,
    pub name: &'a str,
    pub rect: PlotRect,
    pub transform: Transform,
}

/// One text op, for hosts to lay out with whatever fonts they have.
pub(crate) struct ChartText<'a> {
    pub object_id: u32,
    pub text: &'a str,
    pub x: f64,
    pub baseline_y: f64,
    pub width: f64,
    pub font: PlotFont,
    pub color: &'a str,
}

/// The chart primitive for `space`, with at most `budget` parts.
pub(crate) fn chart_primitive(
    frame: ChartFrame<'_>,
    space: &ChartSpace,
    budget: usize,
    text: &mut dyn FnMut(ChartText<'_>) -> Result<Primitive, RenderError>,
) -> Result<Primitive, RenderError> {
    let labels = ChartLabels::new(space);
    let chart = plot_model(space, &labels);
    let mut primitives = Vec::new();
    let mut sink = ChartSink {
        primitives: &mut primitives,
        object_id: frame.object_id,
        text,
        error: None,
        remaining: budget,
    };
    plot_chart_into(&chart, frame.rect, &mut sink);
    if let Some(error) = sink.error {
        return Err(error);
    }
    Ok(Primitive::Chart {
        object_id: frame.object_id,
        shape_id: frame.shape_id.map(str::to_owned),
        name: frame.name.to_owned(),
        x: frame.rect.x as f32,
        y: frame.rect.y as f32,
        w: frame.rect.w as f32,
        h: frame.rect.h as f32,
        label: chart_aria_label(&chart),
        primitives,
        transform: frame.transform,
    })
}

/// Value labels, owned so the plot model can borrow them.
struct ChartLabels {
    groups: Vec<Vec<Vec<(usize, String)>>>,
}

impl ChartLabels {
    fn new(space: &ChartSpace) -> Self {
        let mut budget = MAX_DATA_LABELS;
        Self {
            groups: space
                .plot_groups
                .iter()
                .map(|group| group_labels(group, &mut budget))
                .collect(),
        }
    }

    fn get(&self, group: usize, series: usize) -> Option<&[(usize, String)]> {
        let labels = self.groups.get(group)?.get(series)?;
        (!labels.is_empty()).then_some(labels.as_slice())
    }
}

fn group_labels(group: &ChartPlotGroup, budget: &mut usize) -> Vec<Vec<(usize, String)>> {
    let group_point_labels = indexed_point_labels(group.data_labels.as_ref());
    group
        .series
        .iter()
        .map(|series| {
            if group.data_labels.is_none() && series.data_labels.is_none() {
                return Vec::new();
            }
            let series_point_labels = indexed_point_labels(series.data_labels.as_ref());
            let mut labels = Vec::new();
            for index in 0..series.values.len() {
                if *budget == 0 {
                    break;
                }
                let label = compose_data_label(
                    group,
                    series,
                    group_point_labels.get(&index).copied(),
                    series_point_labels.get(&index).copied(),
                    index,
                );
                if let Some(label) = label {
                    *budget -= 1;
                    labels.push((index, label));
                }
            }
            labels
        })
        .collect()
}

fn indexed_point_labels(labels: Option<&ChartDataLabels>) -> HashMap<usize, &ChartPointLabel> {
    labels
        .and_then(|labels| labels.points.as_deref())
        .into_iter()
        .flatten()
        .filter_map(|labels| data_label_index(labels.index).map(|index| (index, labels)))
        .fold(HashMap::new(), |mut labels, (index, value)| {
            labels.entry(index).or_insert(value);
            labels
        })
}

fn compose_data_label(
    group: &ChartPlotGroup,
    series: &ChartSeries,
    group_point: Option<&ChartPointLabel>,
    series_point: Option<&ChartPointLabel>,
    index: usize,
) -> Option<String> {
    let series_labels = series.data_labels.as_ref();
    let group_labels = group.data_labels.as_ref();
    let levels = [
        series_point.map(|point| &point.labels),
        series_labels,
        group_point.map(|point| &point.labels),
        group_labels,
    ];
    if levels.iter().all(Option::is_none) {
        return None;
    }
    if label_flag(&levels, |labels| labels.delete).unwrap_or(false) {
        return None;
    }
    if let Some(text) = series_point
        .and_then(|point| point.text.as_ref())
        .or_else(|| group_point.and_then(|point| point.text.as_ref()))
    {
        return Some(text.clone());
    }
    let switch =
        |read: fn(&ChartDataLabels) -> Option<bool>| label_flag(&levels, read).unwrap_or(false);
    let separator = label_text(&levels, |labels| labels.separator.as_deref()).unwrap_or(", ");
    let mut parts = Vec::new();
    if switch(|labels| labels.show_series_name)
        && let Some(name) = &series.name
    {
        parts.push(name.clone());
    }
    if switch(|labels| labels.show_category_name) {
        parts.push(
            series
                .categories
                .get(index)
                .cloned()
                .unwrap_or_else(|| (index + 1).to_string()),
        );
    }
    let value = series.values.get(index).copied().unwrap_or_default();
    if switch(|labels| labels.show_value) {
        parts.push(format_number(value));
    }
    if switch(|labels| labels.show_percent) {
        let total = data_label_total(group, series, index);
        parts.push(format_percent(if total > 0.0 {
            value / total
        } else {
            0.0
        }));
    }
    (!parts.is_empty()).then(|| parts.join(separator))
}

fn label_flag(
    levels: &[Option<&ChartDataLabels>],
    read: fn(&ChartDataLabels) -> Option<bool>,
) -> Option<bool> {
    levels
        .iter()
        .copied()
        .find_map(|labels| labels.and_then(read))
}

fn label_text<'a>(
    levels: &[Option<&'a ChartDataLabels>],
    read: fn(&'a ChartDataLabels) -> Option<&'a str>,
) -> Option<&'a str> {
    levels
        .iter()
        .copied()
        .find_map(|labels| labels.and_then(read))
}

fn data_label_total(group: &ChartPlotGroup, series: &ChartSeries, index: usize) -> f64 {
    if matches!(
        group.chart_type.as_deref(),
        Some("pie" | "doughnut" | "ofPie")
    ) {
        series
            .values
            .iter()
            .map(|value| value.abs())
            .filter(|value| value.is_finite())
            .sum()
    } else {
        group
            .series
            .iter()
            .filter_map(|series| series.values.get(index))
            .map(|value| value.abs())
            .filter(|value| value.is_finite())
            .sum()
    }
}

fn format_percent(value: f64) -> String {
    format!("{}%", format_number(value * 100.0))
}

fn data_label_index(index: Option<f64>) -> Option<usize> {
    let index = index?;
    (index.is_finite() && index >= 0.0 && index.fract() == 0.0 && index <= f64::from(u32::MAX))
        .then_some(index as usize)
}

fn plot_model<'a>(space: &'a ChartSpace, labels: &'a ChartLabels) -> PlotChart<'a> {
    let mut chart = PlotChart::from(space);
    for (group_index, group) in chart.plot_groups.iter_mut().enumerate() {
        for (series_index, series) in group.series.iter_mut().enumerate() {
            if let Some(labels) = labels.get(group_index, series_index) {
                series.points = labelled_points(&series.points, labels);
            }
        }
    }
    chart
}

/// One labelled point per label, carrying whatever colour and marker the
/// chart part already resolved for that index, ahead of the original points
/// so the geometry's first-match lookup finds the labelled one.
fn labelled_points<'a>(
    points: &[PlotPoint<'a>],
    labels: &'a [(usize, String)],
) -> Vec<PlotPoint<'a>> {
    let wildcard = points.iter().position(|point| point.index.is_none());
    let mut indexed = points
        .iter()
        .enumerate()
        .filter_map(|(position, point)| point.index.map(|index| (index, position)))
        .collect::<Vec<_>>();
    indexed.sort_unstable();
    indexed.dedup_by_key(|(index, _)| *index);
    let mut output = Vec::with_capacity(labels.len().saturating_add(points.len()));
    for (index, label) in labels {
        let exact = indexed
            .binary_search_by_key(index, |(key, _)| *key)
            .ok()
            .map(|slot| indexed[slot].1);
        let source = match (wildcard, exact) {
            (Some(wildcard), Some(exact)) => Some(wildcard.min(exact)),
            (wildcard, exact) => wildcard.or(exact),
        };
        output.push(PlotPoint {
            index: Some(*index),
            label: Some(label.as_str()),
            ..source.map(|position| points[position]).unwrap_or_default()
        });
    }
    output.extend_from_slice(points);
    output
}

/// Translates plot ops into slide primitives as they are emitted. Chart parts
/// are anonymous: the chart primitive around them carries the identity.
struct ChartSink<'a> {
    primitives: &'a mut Vec<Primitive>,
    object_id: u32,
    text: &'a mut dyn FnMut(ChartText<'_>) -> Result<Primitive, RenderError>,
    error: Option<RenderError>,
    remaining: usize,
}

impl PlotSink for ChartSink<'_> {
    fn push_op(&mut self, op: PlotOp) {
        if self.remaining == 0 || self.error.is_some() {
            return;
        }
        self.remaining -= 1;
        let primitive = match op {
            PlotOp::Rect { x, y, w, h, fill } => self.shape(
                x,
                y,
                w,
                h,
                "rect",
                unit_rectangle(),
                Some(Paint::Solid { color: fill }),
                None,
            ),
            PlotOp::Line {
                x1,
                y1,
                x2,
                y2,
                color,
                width,
            } => {
                let (x, w) = (x1.min(x2), (x2 - x1).abs());
                let (y, h) = (y1.min(y2), (y2 - y1).abs());
                self.shape(
                    x,
                    y,
                    w,
                    h,
                    "line",
                    vec![
                        GeometryPathCommand::Move {
                            x: fraction(x1, x, w),
                            y: fraction(y1, y, h),
                        },
                        GeometryPathCommand::Line {
                            x: fraction(x2, x, w),
                            y: fraction(y2, y, h),
                        },
                    ],
                    None,
                    Some(Stroke {
                        color,
                        width: width as f32,
                        dashed: false,
                    }),
                )
            }
            PlotOp::Path {
                x,
                y,
                w,
                h,
                commands,
                fill,
                stroke,
            } => self.shape(
                x,
                y,
                w,
                h,
                "custom",
                commands
                    .into_iter()
                    .map(|command| normalize(command, x, y, w, h))
                    .collect(),
                Some(Paint::Solid { color: fill }),
                stroke.map(|stroke| Stroke {
                    color: stroke.color,
                    width: stroke.width as f32,
                    dashed: false,
                }),
            ),
            PlotOp::Text {
                text,
                x,
                baseline_y,
                width,
                font,
                color,
            } => {
                let request = ChartText {
                    object_id: self.object_id,
                    text: &text,
                    x,
                    baseline_y,
                    width,
                    font,
                    color: &color,
                };
                match (self.text)(request) {
                    Ok(primitive) => primitive,
                    Err(error) => {
                        self.error = Some(error);
                        return;
                    }
                }
            }
        };
        self.primitives.push(primitive);
    }
}

impl ChartSink<'_> {
    #[allow(clippy::too_many_arguments)]
    fn shape(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        geometry: &str,
        path: Vec<GeometryPathCommand>,
        fill: Option<Paint>,
        stroke: Option<Stroke>,
    ) -> Primitive {
        Primitive::Shape {
            object_id: self.object_id,
            shape_id: None,
            name: String::new(),
            x: x as f32,
            y: y as f32,
            w: w as f32,
            h: h as f32,
            geometry: geometry.to_owned(),
            path,
            adjust_values: Default::default(),
            fill,
            stroke,
            transform: Transform::default(),
        }
    }
}

fn unit_rectangle() -> Vec<GeometryPathCommand> {
    vec![
        GeometryPathCommand::Move { x: 0.0, y: 0.0 },
        GeometryPathCommand::Line { x: 1.0, y: 0.0 },
        GeometryPathCommand::Line { x: 1.0, y: 1.0 },
        GeometryPathCommand::Line { x: 0.0, y: 1.0 },
        GeometryPathCommand::Close,
    ]
}

/// Where `value` sits inside `[origin, origin + extent]`. A zero extent has no
/// inside, and painting `origin + fraction * 0` puts every point on the edge.
fn fraction(value: f64, origin: f64, extent: f64) -> f64 {
    if extent == 0.0 {
        0.0
    } else {
        (value - origin) / extent
    }
}

fn normalize(command: GeometryPathCommand, x: f64, y: f64, w: f64, h: f64) -> GeometryPathCommand {
    let px = |value| fraction(value, x, w);
    let py = |value| fraction(value, y, h);
    match command {
        GeometryPathCommand::Move { x, y } => GeometryPathCommand::Move { x: px(x), y: py(y) },
        GeometryPathCommand::Line { x, y } => GeometryPathCommand::Line { x: px(x), y: py(y) },
        GeometryPathCommand::Quad { cpx, cpy, x, y } => GeometryPathCommand::Quad {
            cpx: px(cpx),
            cpy: py(cpy),
            x: px(x),
            y: py(y),
        },
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => GeometryPathCommand::Cubic {
            cp1x: px(cp1x),
            cp1y: py(cp1y),
            cp2x: px(cp2x),
            cp2y: py(cp2y),
            x: px(x),
            y: py(y),
        },
        GeometryPathCommand::Close => GeometryPathCommand::Close,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_parse::{ChartAxes, ChartAxis, ChartLegend, ChartMarker, ChartPoint, ChartSeries};

    fn series(name: &str, values: &[f64], color: &str) -> ChartSeries {
        ChartSeries {
            name: Some(name.to_owned()),
            categories: ["Q1", "Q2", "Q3"]
                .iter()
                .take(values.len())
                .map(|value| (*value).to_owned())
                .collect(),
            values: values.to_vec(),
            color: color.to_owned(),
            index: None,
            order: None,
            category_formula: None,
            value_formula: None,
            axis_ids: None,
            points: None,
            grouping: None,
            marker: None,
            smooth: None,
            data_labels: None,
        }
    }

    fn value_labels() -> ChartDataLabels {
        ChartDataLabels {
            show_value: Some(true),
            ..ChartDataLabels::default()
        }
    }

    fn group(chart_type: &str, series: Vec<ChartSeries>) -> ChartPlotGroup {
        ChartPlotGroup {
            chart_type: Some(chart_type.to_owned()),
            grouping: Some("clustered".to_owned()),
            overlap: None,
            gap_width: None,
            axis_ids: Vec::new(),
            series,
            vary_colors: false,
            first_slice_angle: None,
            hole_size: None,
            data_labels: None,
        }
    }

    fn space(chart_type: &str, groups: Vec<ChartPlotGroup>) -> ChartSpace {
        ChartSpace {
            chart_type: chart_type.to_owned(),
            title: Some("Revenue".to_owned()),
            legend: Some(ChartLegend {
                position: Some("right".to_owned()),
                visible: true,
            }),
            series: groups
                .iter()
                .flat_map(|group| group.series.iter())
                .cloned()
                .collect(),
            axes: None,
            plot_groups: groups,
            axis_list: None,
        }
    }

    fn frame(name: &str) -> ChartFrame<'_> {
        ChartFrame {
            object_id: 7,
            shape_id: Some("slide:1"),
            name,
            rect: PlotRect {
                x: 10.0,
                y: 20.0,
                w: 320.0,
                h: 240.0,
            },
            transform: Transform::default(),
        }
    }

    /// Plots `space` with a text callback that needs no fonts.
    fn plot(space: &ChartSpace) -> Primitive {
        chart_primitive(frame("Chart 1"), space, 100_000, &mut |text| {
            Ok(Primitive::TextBox {
                object_id: text.object_id,
                shape_id: None,
                story_id: None,
                x: text.x as f32,
                y: text.baseline_y as f32,
                w: text.width as f32,
                h: text.font.size_px as f32,
                anchor: crate::TextAnchor::Top,
                paragraphs: vec![crate::TextParagraph {
                    align: None,
                    level: 0,
                    runs: vec![crate::TextRun {
                        text: text.text.to_owned(),
                        font_family: text.font.family.to_owned(),
                        font_size_pt: text.font.size_px as f32,
                        bold: text.font.weight >= 600,
                        italic: false,
                        underline: false,
                        color: text.color.to_owned(),
                    }],
                }],
                lines: Vec::new(),
                overflow: false,
                transform: Transform::default(),
            })
        })
        .expect("chart plots")
    }

    fn parts(chart: &Primitive) -> &[Primitive] {
        let Primitive::Chart { primitives, .. } = chart else {
            panic!("expected a chart primitive");
        };
        primitives
    }

    fn label(chart: &Primitive) -> &str {
        let Primitive::Chart { label, .. } = chart else {
            panic!("expected a chart primitive");
        };
        label
    }

    fn texts(chart: &Primitive) -> Vec<String> {
        parts(chart)
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::TextBox { paragraphs, .. } => Some(paragraphs[0].runs[0].text.clone()),
                _ => None,
            })
            .collect()
    }

    fn geometries(chart: &Primitive) -> Vec<&str> {
        parts(chart)
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape { geometry, .. } => Some(geometry.as_str()),
                _ => None,
            })
            .collect()
    }

    fn fills(chart: &Primitive) -> Vec<&str> {
        parts(chart)
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape {
                    fill: Some(Paint::Solid { color }),
                    ..
                } => Some(color.as_str()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn every_parsed_chart_type_plots_into_primitives() {
        for chart_type in [
            "column", "bar", "line", "pie", "doughnut", "area", "scatter", "radar", "stock",
            "bubble", "surface", "ofPie",
        ] {
            let space = space(
                chart_type,
                vec![group(
                    chart_type,
                    vec![series("North", &[3.0, 1.0], "#112233")],
                )],
            );
            let chart = plot(&space);
            let parts = parts(&chart);
            assert!(!parts.is_empty(), "{chart_type} drew nothing");
            assert!(
                fills(&chart).contains(&"#112233") || geometries(&chart).contains(&"custom"),
                "{chart_type} lost its series colour"
            );
            assert!(
                texts(&chart).contains(&"Revenue".to_owned()),
                "{chart_type} lost its title"
            );
            assert!(
                texts(&chart)
                    .iter()
                    .any(|text| text == "North" || text == "Q1"),
                "{chart_type} lost its legend"
            );
        }
    }

    #[test]
    fn wedge_families_draw_closed_paths_and_flat_families_draw_rectangles() {
        for chart_type in ["pie", "doughnut", "ofPie"] {
            let space = space(
                chart_type,
                vec![group(
                    chart_type,
                    vec![series("Share", &[3.0, 1.0], "#112233")],
                )],
            );
            let paths = parts(&plot(&space))
                .iter()
                .filter(|primitive| {
                    matches!(primitive, Primitive::Shape { geometry, path, .. }
                        if geometry == "custom"
                            && matches!(path.last(), Some(GeometryPathCommand::Close)))
                })
                .count();
            assert_eq!(paths, 2, "{chart_type}");
        }
        for chart_type in ["column", "bar"] {
            let space = space(
                chart_type,
                vec![group(
                    chart_type,
                    vec![series("North", &[3.0, 1.0], "#112233")],
                )],
            );
            let chart = plot(&space);
            assert!(geometries(&chart).contains(&"rect"));
            assert!(geometries(&chart).contains(&"line"));
        }
    }

    #[test]
    fn a_legend_plots_in_every_position() {
        for position in ["left", "right", "top", "bottom"] {
            let mut space = space(
                "column",
                vec![group(
                    "column",
                    vec![series("North", &[3.0, 1.0], "#112233")],
                )],
            );
            space.legend = Some(ChartLegend {
                position: Some(position.to_owned()),
                visible: true,
            });
            assert!(
                texts(&plot(&space)).contains(&"North".to_owned()),
                "{position}"
            );
        }
        let mut hidden = space(
            "column",
            vec![group(
                "column",
                vec![series("North", &[3.0, 1.0], "#112233")],
            )],
        );
        hidden.legend = Some(ChartLegend {
            position: None,
            visible: false,
        });
        assert!(!texts(&plot(&hidden)).contains(&"North".to_owned()));
    }

    #[test]
    fn combo_groups_plot_every_group_against_the_shared_axis() {
        let mut space = space(
            "column",
            vec![
                group("column", vec![series("Revenue", &[5.0, 9.0], "#112233")]),
                group("line", vec![series("Trend", &[4.0, 8.0], "#445566")]),
            ],
        );
        space.axes = Some(ChartAxes {
            category: None,
            value: Some(ChartAxis {
                id: None,
                title: Some("Millions".to_owned()),
                min: Some(0.0),
                max: Some(10.0),
                labels: None,
                axis_type: "value".to_owned(),
                position: Some("left".to_owned()),
                cross_axis_id: None,
                crosses: None,
                crosses_at: None,
                major_unit: None,
                minor_unit: None,
                logarithmic_base: None,
                reversed: false,
                number_format: None,
                major_tick_mark: None,
                minor_tick_mark: None,
                tick_label_position: None,
                hidden: false,
            }),
        });
        let chart = plot(&space);
        assert!(fills(&chart).contains(&"#112233"));
        assert!(parts(&chart).iter().any(
            |primitive| matches!(primitive, Primitive::Shape { geometry, stroke: Some(stroke), .. }
                    if geometry == "line" && stroke.color == "#445566")
        ));
        assert!(texts(&chart).contains(&"10".to_owned()));
        assert_eq!(
            label(&chart),
            "Revenue, combo chart, 2 series, 2 categories"
        );
    }

    #[test]
    fn per_point_colours_markers_and_data_labels_reach_the_primitives() {
        let mut point_series = series("North", &[3.0, 1.0, 2.0], "#112233");
        point_series.marker = Some(ChartMarker {
            symbol: Some("circle".to_owned()),
            size: Some(11.0),
        });
        point_series.points = Some(vec![ChartPoint {
            index: Some(1.0),
            explosion: None,
            color: "#AABBCC".to_owned(),
        }]);
        let mut group = group("line", vec![point_series]);
        group.data_labels = Some(value_labels());
        let chart = plot(&space("line", vec![group]));

        assert!(fills(&chart).contains(&"#AABBCC"));
        assert!(parts(&chart).iter().any(
            |primitive| matches!(primitive, Primitive::Shape { geometry, w, .. }
                if geometry == "rect" && (*w - 11.0).abs() < 0.001)
        ));
        for value in ["3", "1", "2"] {
            assert!(texts(&chart).contains(&value.to_owned()), "{value}");
        }
    }

    #[test]
    fn the_label_budget_is_chart_wide_rather_than_per_series() {
        let values = (0..3_000).map(|value| value as f64).collect::<Vec<_>>();
        let mut group = group(
            "line",
            vec![
                series("First", &values, "#112233"),
                series("Second", &values, "#445566"),
            ],
        );
        group.data_labels = Some(value_labels());
        let space = space("line", vec![group]);
        let labels = ChartLabels::new(&space);
        assert_eq!(labels.get(0, 0).unwrap().len(), 3_000);
        assert_eq!(labels.get(0, 1).unwrap().len(), MAX_DATA_LABELS - 3_000);
    }

    #[test]
    fn series_and_point_label_overrides_do_not_leak() {
        let mut shown = series("Shown", &[11.0, 12.0], "#112233");
        shown.data_labels = Some(ChartDataLabels {
            show_value: Some(true),
            points: Some(vec![ChartPointLabel {
                index: Some(1.0),
                text: None,
                labels: ChartDataLabels {
                    delete: Some(true),
                    ..ChartDataLabels::default()
                },
            }]),
            ..ChartDataLabels::default()
        });
        let inherited = series("Inherited", &[21.0, 22.0], "#445566");
        let mut deleted = series("Deleted", &[31.0, 32.0], "#778899");
        deleted.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            ..ChartDataLabels::default()
        });
        let group = group("line", vec![shown, inherited, deleted]);
        let space = space("line", vec![group]);

        let labels = ChartLabels::new(&space);

        assert_eq!(labels.get(0, 0), Some([(0, "11".to_owned())].as_slice()));
        assert!(labels.get(0, 1).is_none());
        assert!(labels.get(0, 2).is_none());
    }

    #[test]
    fn series_delete_overrides_group_label_defaults() {
        let inherited = series("Inherited", &[41.0, 42.0], "#112233");
        let mut deleted = series("Deleted", &[51.0], "#445566");
        deleted.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            ..ChartDataLabels::default()
        });
        let mut group = group("line", vec![inherited, deleted]);
        group.data_labels = Some(ChartDataLabels {
            show_value: Some(true),
            points: Some(vec![ChartPointLabel {
                index: Some(1.0),
                text: None,
                labels: ChartDataLabels {
                    delete: Some(true),
                    ..ChartDataLabels::default()
                },
            }]),
            ..ChartDataLabels::default()
        });
        let space = space("line", vec![group]);

        let labels = ChartLabels::new(&space);

        assert_eq!(labels.get(0, 0), Some([(0, "41".to_owned())].as_slice()));
        assert!(labels.get(0, 1).is_none());
    }

    #[test]
    fn delete_is_inherited_until_a_lower_scope_overrides_it() {
        let mut inherited = series("Inherited", &[11.0, 12.0], "#112233");
        inherited.data_labels = Some(ChartDataLabels {
            show_category_name: Some(false),
            ..ChartDataLabels::default()
        });
        let mut point_restored = series("Point", &[21.0, 22.0], "#445566");
        point_restored.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            points: Some(vec![ChartPointLabel {
                index: Some(1.0),
                text: None,
                labels: ChartDataLabels {
                    delete: Some(false),
                    ..ChartDataLabels::default()
                },
            }]),
            ..ChartDataLabels::default()
        });
        let mut series_restored = series("Series", &[31.0, 32.0], "#778899");
        series_restored.data_labels = Some(ChartDataLabels {
            delete: Some(false),
            ..ChartDataLabels::default()
        });
        let mut group = group("line", vec![inherited, point_restored, series_restored]);
        group.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            show_value: Some(true),
            ..ChartDataLabels::default()
        });
        let space = space("line", vec![group]);

        let labels = ChartLabels::new(&space);

        assert!(labels.get(0, 0).is_none());
        assert_eq!(labels.get(0, 1), Some([(1, "22".to_owned())].as_slice()));
        assert_eq!(
            labels.get(0, 2),
            Some([(0, "31".to_owned()), (1, "32".to_owned())].as_slice())
        );
    }

    #[test]
    fn data_labels_keep_the_colour_the_chart_part_gave_each_point() {
        let mut point_series = series("Share", &[3.0, 1.0], "#112233");
        point_series.points = Some(vec![ChartPoint {
            index: Some(0.0),
            explosion: None,
            color: "#AABBCC".to_owned(),
        }]);
        let mut group = group("pie", vec![point_series]);
        group.data_labels = Some(value_labels());
        let chart = plot(&space("pie", vec![group]));

        assert!(fills(&chart).contains(&"#AABBCC"));
        assert!(texts(&chart).contains(&"3".to_owned()));
    }

    #[test]
    fn the_part_budget_bounds_a_chart_with_more_data_than_it_can_draw() {
        let values = (0..50_000).map(|value| value as f64).collect::<Vec<_>>();
        let space = space(
            "line",
            vec![group("line", vec![series("Wide", &values, "#112233")])],
        );
        assert_eq!(parts(&plot(&space)).len(), 100_000);
        let chart = chart_primitive(frame("Wide"), &space, 64, &mut |text| {
            Ok(Primitive::Placeholder {
                object_id: text.object_id,
                shape_id: None,
                name: text.text.to_owned(),
                x: 0.0,
                y: 0.0,
                w: 0.0,
                h: 0.0,
                label: None,
                transform: Transform::default(),
            })
        })
        .expect("chart plots");
        assert_eq!(parts(&chart).len(), 64);
    }

    #[test]
    fn a_degenerate_frame_keeps_every_coordinate_finite() {
        let space = space(
            "column",
            vec![group(
                "column",
                vec![series("North", &[3.0, 1.0], "#112233")],
            )],
        );
        for rect in [
            PlotRect {
                x: f64::NAN,
                y: 0.0,
                w: 300.0,
                h: 200.0,
            },
            PlotRect {
                x: 0.0,
                y: 0.0,
                w: f64::INFINITY,
                h: 0.0,
            },
            PlotRect::default(),
        ] {
            let chart = chart_primitive(
                ChartFrame {
                    rect,
                    ..frame("Degenerate")
                },
                &space,
                100_000,
                &mut |_| {
                    Ok(Primitive::Placeholder {
                        object_id: 0,
                        shape_id: None,
                        name: String::new(),
                        x: 0.0,
                        y: 0.0,
                        w: 0.0,
                        h: 0.0,
                        label: None,
                        transform: Transform::default(),
                    })
                },
            )
            .expect("chart plots");
            for primitive in parts(&chart) {
                let Primitive::Shape {
                    x, y, w, h, path, ..
                } = primitive
                else {
                    continue;
                };
                assert!(
                    x.is_finite() && y.is_finite() && w.is_finite() && h.is_finite(),
                    "{rect:?}"
                );
                for command in path {
                    let (x, y) = match command {
                        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                            (*x, *y)
                        }
                        _ => continue,
                    };
                    assert!(x.is_finite() && y.is_finite(), "{rect:?}");
                }
            }
        }
    }
}
