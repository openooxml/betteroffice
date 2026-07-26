//! Format-neutral chart geometry: a chart plus a rectangle in, an ordered
//! [`PlotOp`] stream out. Hosts translate the ops into their own primitives.

use crate::GeometryPathCommand;

use super::model::ChartSpace;

pub const CHART_AXIS_COLOR: &str = "#666666";
pub const CHART_GRID_COLOR: &str = "#D9D9D9";
pub const CHART_TEXT_COLOR: &str = "#222222";
pub const CHART_BACKGROUND_COLOR: &str = "#FFFFFF";
pub const CHART_SERIES_COLORS: [&str; 8] = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E",
];
pub const CHART_LABEL_FONT: PlotFont = PlotFont {
    weight: 400,
    size_px: 10.0,
    family: "Calibri, sans-serif",
};
pub const CHART_TITLE_FONT: PlotFont = PlotFont {
    weight: 600,
    size_px: 13.0,
    family: "Calibri, sans-serif",
};

/// Hard ceiling on the ops one chart may emit, whatever its data length.
pub const MAX_PLOT_OPS: usize = 100_000;

const MAX_LABEL_CHARS: usize = 120;
const MAX_LEGEND_ENTRIES: usize = 8;

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PlotFont {
    pub weight: u16,
    pub size_px: f64,
    pub family: &'static str,
}

impl PlotFont {
    /// CSS `font` shorthand, for hosts that paint through a browser.
    pub fn css(&self) -> String {
        format!("{} {}px {}", self.weight, self.size_px, self.family)
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotRect {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

#[derive(Clone, Debug, PartialEq)]
pub struct PlotStroke {
    pub color: String,
    pub width: f64,
}

/// One draw instruction, in the same coordinate space as the input rectangle.
#[derive(Clone, Debug, PartialEq)]
pub enum PlotOp {
    Rect {
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: String,
    },
    Text {
        text: String,
        x: f64,
        baseline_y: f64,
        width: f64,
        font: PlotFont,
        color: String,
    },
    Line {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        color: String,
        width: f64,
    },
    Path {
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        commands: Vec<GeometryPathCommand>,
        fill: String,
        stroke: Option<PlotStroke>,
    },
}

/// Receives plot ops in back-to-front order as they are produced.
pub trait PlotSink {
    fn push_op(&mut self, op: PlotOp);
}

impl PlotSink for Vec<PlotOp> {
    fn push_op(&mut self, op: PlotOp) {
        self.push(op);
    }
}

/// What the geometry reads off a chart. Distinct from [`ChartSpace`], which is
/// the parse-fidelity model: hosts may inject per-point values and labels that
/// never appear in the chart part. Borrows its host's strings and data.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotChart<'a> {
    pub chart_type: &'a str,
    pub title: Option<&'a str>,
    pub legend: Option<PlotLegend<'a>>,
    pub value_axis: Option<PlotAxisRange>,
    pub series: Vec<PlotSeries<'a>>,
    pub plot_groups: Vec<PlotGroup<'a>>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotLegend<'a> {
    pub position: Option<&'a str>,
    pub visible: Option<bool>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotAxisRange {
    pub min: Option<f64>,
    pub max: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotGroup<'a> {
    pub chart_type: Option<&'a str>,
    pub grouping: Option<&'a str>,
    pub series: Vec<PlotSeries<'a>>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotSeries<'a> {
    pub name: Option<&'a str>,
    pub categories: &'a [String],
    pub values: &'a [f64],
    pub color: Option<&'a str>,
    pub points: Vec<PlotPoint<'a>>,
    pub grouping: Option<&'a str>,
    pub marker: Option<PlotMarker>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotPoint<'a> {
    pub index: Option<usize>,
    pub value: Option<f64>,
    pub color: Option<&'a str>,
    pub marker: Option<PlotMarker>,
    pub label: Option<&'a str>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotMarker {
    pub size: Option<f64>,
}

impl<'a> From<&'a ChartSpace> for PlotChart<'a> {
    fn from(space: &'a ChartSpace) -> Self {
        Self {
            chart_type: &space.chart_type,
            title: space.title.as_deref(),
            legend: space.legend.as_ref().map(|legend| PlotLegend {
                position: legend.position.as_deref(),
                visible: Some(legend.visible),
            }),
            value_axis: space
                .axes
                .as_ref()
                .and_then(|axes| axes.value.as_ref())
                .map(|axis| PlotAxisRange {
                    min: axis.min,
                    max: axis.max,
                }),
            series: space.series.iter().map(plot_series_from_model).collect(),
            plot_groups: space
                .plot_groups
                .iter()
                .map(|group| PlotGroup {
                    chart_type: group.chart_type.as_deref(),
                    grouping: group.grouping.as_deref(),
                    series: group.series.iter().map(plot_series_from_model).collect(),
                })
                .collect(),
        }
    }
}

fn plot_series_from_model(series: &super::model::ChartSeries) -> PlotSeries<'_> {
    PlotSeries {
        name: series.name.as_deref(),
        categories: &series.categories,
        values: &series.values,
        color: Some(&series.color),
        points: series
            .points
            .iter()
            .flatten()
            .map(|point| PlotPoint {
                index: point.index.map(|index| index as usize),
                value: None,
                color: Some(&point.color),
                marker: None,
                label: None,
            })
            .collect(),
        grouping: series.grouping.as_deref(),
        marker: series
            .marker
            .as_ref()
            .map(|marker| PlotMarker { size: marker.size }),
    }
}

/// Draw ops for `chart` inside `rect`, back to front, into `sink`.
pub fn plot_chart_into<S: PlotSink + ?Sized>(chart: &PlotChart<'_>, rect: PlotRect, sink: &mut S) {
    let PlotRect {
        x,
        y,
        w: width,
        h: height,
    } = rect;
    let ops = &mut Emitter {
        sink,
        remaining: MAX_PLOT_OPS,
    };

    push_rect(ops, x, y, width, height, CHART_BACKGROUND_COLOR);

    let title_h = if let Some(title) = chart.title.filter(|s| !s.is_empty()) {
        push_text(
            ops,
            title,
            x + 8.0,
            y + 18.0,
            (width - 16.0).max(0.0),
            CHART_TITLE_FONT,
        );
        28.0
    } else {
        10.0
    };

    let legend_position = chart
        .legend
        .as_ref()
        .and_then(|legend| legend.position)
        .unwrap_or("right");
    let legend_w = if has_legend(chart) { 104.0 } else { 8.0 };
    let plot_x = if legend_position == "left" {
        x + legend_w + 42.0
    } else {
        x + 42.0
    };
    let plot = PlotArea {
        x: plot_x,
        y: y + title_h,
        w: (width - 42.0 - legend_w - 10.0).max(24.0),
        h: (height - title_h - 34.0).max(24.0),
    };

    if chart.plot_groups.is_empty() {
        emit_family(
            ops,
            PlotFamily {
                chart_type: chart.chart_type,
                series: &chart.series,
                value_axis: chart.value_axis,
            },
            plot,
            x,
            y + title_h,
            width,
            height - title_h,
        );
    } else {
        for group in &chart.plot_groups {
            emit_family(
                ops,
                PlotFamily {
                    chart_type: group.chart_type.unwrap_or(chart.chart_type),
                    series: &group.series,
                    value_axis: chart.value_axis,
                },
                plot,
                x,
                y + title_h,
                width,
                height - title_h,
            );
        }
    }

    let legend_x = if legend_position == "left" {
        x + 6.0
    } else {
        x + width - legend_w + 6.0
    };
    emit_legend(ops, chart, legend_x, y + title_h + 8.0, legend_w - 12.0);
}

/// Draw ops for `chart` inside `rect`, collected into one vector.
pub fn plot_chart(chart: &PlotChart<'_>, rect: PlotRect) -> Vec<PlotOp> {
    let mut ops = Vec::new();
    plot_chart_into(chart, rect, &mut ops);
    ops
}

/// Screen-reader summary of a chart.
pub fn chart_aria_label(chart: &PlotChart<'_>) -> String {
    let kind = if chart.plot_groups.len() > 1 {
        "combo chart"
    } else {
        match chart.chart_type {
            "bar" => "bar chart",
            "line" => "line chart",
            "pie" => "pie chart",
            "doughnut" => "doughnut chart",
            _ => "column chart",
        }
    };
    let title = chart
        .title
        .filter(|s| !s.is_empty())
        .unwrap_or("Untitled chart");
    let series_count = if chart.series.is_empty() {
        chart
            .plot_groups
            .iter()
            .map(|group| group.series.len())
            .sum()
    } else {
        chart.series.len()
    };
    let category_count = if chart.series.is_empty() {
        chart
            .plot_groups
            .iter()
            .flat_map(|group| group.series.iter())
            .map(series_length)
            .max()
            .unwrap_or(0)
    } else {
        category_count(&chart.series)
    };
    format!("{title}, {kind}, {series_count} series, {category_count} categories")
}

/// Op sink plus the remaining chart-wide op budget.
struct Emitter<'s, S: PlotSink + ?Sized> {
    sink: &'s mut S,
    remaining: usize,
}

impl<S: PlotSink + ?Sized> Emitter<'_, S> {
    fn push(&mut self, op: PlotOp) {
        if self.remaining == 0 {
            return;
        }
        self.remaining -= 1;
        self.sink.push_op(op);
    }

    fn exhausted(&self) -> bool {
        self.remaining == 0
    }
}

/// The slice of a chart one plot family draws: a combo chart emits one per
/// plot group, all sharing the outer chart's value axis.
#[derive(Clone, Copy)]
struct PlotFamily<'a, 'b> {
    chart_type: &'a str,
    series: &'b [PlotSeries<'a>],
    value_axis: Option<PlotAxisRange>,
}

#[derive(Clone, Copy)]
struct PlotArea {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}

fn emit_family<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_, '_>,
    plot: PlotArea,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    match family.chart_type {
        "pie" | "doughnut" => emit_pie(ops, family, x, y, width, height),
        "line" | "scatter" | "radar" => emit_line(ops, family, plot),
        "bar" => emit_bar(ops, family, plot, true),
        _ => emit_bar(ops, family, plot, false),
    }
}

fn push_rect<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    fill: &str,
) {
    if w <= 0.0 || h <= 0.0 {
        return;
    }
    ops.push(PlotOp::Rect {
        x,
        y,
        w,
        h,
        fill: fill.to_owned(),
    });
}

fn push_text<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    text: &str,
    x: f64,
    baseline_y: f64,
    width: f64,
    font: PlotFont,
) {
    if text.is_empty() || width <= 0.0 {
        return;
    }
    ops.push(PlotOp::Text {
        text: text.chars().take(MAX_LABEL_CHARS).collect(),
        x,
        baseline_y,
        width,
        font,
        color: CHART_TEXT_COLOR.to_owned(),
    });
}

fn push_line<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    x1: f64,
    y1: f64,
    x2: f64,
    y2: f64,
    color: &str,
    width: f64,
) {
    ops.push(PlotOp::Line {
        x1,
        y1,
        x2,
        y2,
        color: color.to_owned(),
        width,
    });
}

fn has_legend(chart: &PlotChart<'_>) -> bool {
    chart
        .legend
        .as_ref()
        .and_then(|legend| legend.visible)
        .unwrap_or(true)
}

fn hex(color: &str) -> String {
    if color.starts_with('#') {
        color.to_owned()
    } else {
        format!("#{color}")
    }
}

fn series_color(series: Option<&PlotSeries<'_>>, index: usize) -> String {
    series
        .and_then(|series| series.color)
        .filter(|color| !color.is_empty())
        .map(hex)
        .unwrap_or_else(|| CHART_SERIES_COLORS[index % CHART_SERIES_COLORS.len()].to_owned())
}

fn series_point<'a, 'p>(series: &'p PlotSeries<'a>, index: usize) -> Option<&'p PlotPoint<'a>> {
    series
        .points
        .iter()
        .find(|point| point.index.unwrap_or(index) == index)
}

fn series_value(series: &PlotSeries<'_>, index: usize) -> f64 {
    series_point(series, index)
        .and_then(|point| point.value)
        .or_else(|| series.values.get(index).copied())
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
}

fn series_point_color(series: &PlotSeries<'_>, point_index: usize, series_index: usize) -> String {
    series_point(series, point_index)
        .and_then(|point| point.color)
        .map(hex)
        .unwrap_or_else(|| series_color(Some(series), series_index))
}

fn series_length(series: &PlotSeries<'_>) -> usize {
    series
        .categories
        .len()
        .max(series.values.len())
        .max(series.points.len())
}

fn category_count(series: &[PlotSeries<'_>]) -> usize {
    series.iter().map(series_length).max().unwrap_or(0)
}

fn category_label(series: &[PlotSeries<'_>], index: usize) -> String {
    series
        .iter()
        .find_map(|series| series.categories.get(index).cloned())
        .unwrap_or_else(|| (index + 1).to_string())
}

fn value_range(family: PlotFamily<'_, '_>) -> (f64, f64) {
    let mut min = 0.0;
    let mut max = 0.0;
    for series in family.series {
        for index in 0..series.values.len().max(series.points.len()) {
            let value = series_value(series, index);
            if value.is_finite() {
                min = f64::min(min, value);
                max = f64::max(max, value);
            }
        }
    }
    if let Some(axis) = family.value_axis.as_ref() {
        if let Some(value) = axis.min.filter(|value| value.is_finite()) {
            min = value;
        }
        if let Some(value) = axis.max.filter(|value| value.is_finite()) {
            max = value;
        }
    }
    if max <= min {
        max = min + 1.0;
    }
    (min, max)
}

fn value_y(plot: PlotArea, value: f64, min: f64, max: f64) -> f64 {
    plot.y + (max - value) / (max - min) * plot.h
}

fn marker_size(series: &PlotSeries<'_>, index: usize) -> f64 {
    series_point(series, index)
        .and_then(|point| point.marker.as_ref())
        .or(series.marker.as_ref())
        .and_then(|marker| marker.size)
        .unwrap_or(4.0)
        .clamp(1.0, 24.0)
}

fn emit_axes<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_, '_>,
    plot: PlotArea,
) {
    let (min, max) = value_range(family);
    for i in 0..=4 {
        let t = i as f64 / 4.0;
        let y = plot.y + t * plot.h;
        push_line(ops, plot.x, y, plot.x + plot.w, y, CHART_GRID_COLOR, 0.5);
        let value = max - t * (max - min);
        push_text(
            ops,
            &format_number(value),
            plot.x - 38.0,
            y + 3.0,
            34.0,
            CHART_LABEL_FONT,
        );
    }
    push_line(
        ops,
        plot.x,
        plot.y,
        plot.x,
        plot.y + plot.h,
        CHART_AXIS_COLOR,
        1.0,
    );
    push_line(
        ops,
        plot.x,
        plot.y + plot.h,
        plot.x + plot.w,
        plot.y + plot.h,
        CHART_AXIS_COLOR,
        1.0,
    );
}

fn emit_bar<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_, '_>,
    plot: PlotArea,
    horizontal: bool,
) {
    let cat_count = category_count(family.series);
    if cat_count == 0 || family.series.is_empty() {
        return;
    }
    emit_axes(ops, family, plot);
    let (min, max) = value_range(family);
    let zero_y = value_y(plot, 0.0_f64.clamp(min, max), min, max);
    let series_count = family.series.len().max(1);
    if horizontal {
        let row_h = plot.h / cat_count as f64;
        let bar_h = (row_h * 0.7 / series_count as f64).max(1.0);
        for cat_idx in 0..cat_count {
            if ops.exhausted() {
                return;
            }
            let label = category_label(family.series, cat_idx);
            push_text(
                ops,
                &label,
                plot.x - 38.0,
                plot.y + row_h * (cat_idx as f64 + 0.55),
                36.0,
                CHART_LABEL_FONT,
            );
            for (ser_idx, series) in family.series.iter().enumerate() {
                let value = series_value(series, cat_idx);
                let ratio = ((value - min) / (max - min)).clamp(0.0, 1.0);
                let bar_w = ratio * plot.w;
                let y = plot.y + row_h * cat_idx as f64 + row_h * 0.15 + bar_h * ser_idx as f64;
                push_rect(
                    ops,
                    plot.x,
                    y,
                    bar_w,
                    bar_h,
                    &series_point_color(series, cat_idx, ser_idx),
                );
            }
        }
    } else {
        let group_w = plot.w / cat_count as f64;
        let bar_w = (group_w * 0.7 / series_count as f64).max(1.0);
        for cat_idx in 0..cat_count {
            if ops.exhausted() {
                return;
            }
            let label = category_label(family.series, cat_idx);
            push_text(
                ops,
                &label,
                plot.x + group_w * cat_idx as f64 + 2.0,
                plot.y + plot.h + 14.0,
                group_w - 4.0,
                CHART_LABEL_FONT,
            );
            for (ser_idx, series) in family.series.iter().enumerate() {
                let value = series_value(series, cat_idx);
                let yv = value_y(plot, value.clamp(min, max), min, max);
                let y0 = zero_y;
                let x = plot.x + group_w * cat_idx as f64 + group_w * 0.15 + bar_w * ser_idx as f64;
                push_rect(
                    ops,
                    x,
                    yv.min(y0),
                    bar_w,
                    (y0 - yv).abs().max(1.0),
                    &series_point_color(series, cat_idx, ser_idx),
                );
            }
        }
    }
}

fn emit_line<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_, '_>,
    plot: PlotArea,
) {
    let cat_count = category_count(family.series);
    if cat_count == 0 || family.series.is_empty() {
        return;
    }
    emit_axes(ops, family, plot);
    let (min, max) = value_range(family);
    let denom = (cat_count.saturating_sub(1)).max(1) as f64;
    for i in 0..cat_count {
        if ops.exhausted() {
            return;
        }
        let label = category_label(family.series, i);
        let x = plot.x + plot.w * i as f64 / denom;
        push_text(
            ops,
            &label,
            x - 16.0,
            plot.y + plot.h + 14.0,
            32.0,
            CHART_LABEL_FONT,
        );
    }
    for (ser_idx, series) in family.series.iter().enumerate() {
        let color = series_color(Some(series), ser_idx);
        let mut prev: Option<(f64, f64)> = None;
        for i in 0..cat_count {
            if ops.exhausted() {
                return;
            }
            let value = series_value(series, i);
            let x = plot.x + plot.w * i as f64 / denom;
            let y = value_y(plot, value.clamp(min, max), min, max);
            if let Some((prev_x, prev_y)) = prev {
                push_line(ops, prev_x, prev_y, x, y, &color, 2.0);
            }
            let size = marker_size(series, i);
            let point_color = series_point_color(series, i, ser_idx);
            push_rect(
                ops,
                x - size / 2.0,
                y - size / 2.0,
                size,
                size,
                &point_color,
            );
            if let Some(label) = series_point(series, i).and_then(|point| point.label) {
                push_text(ops, label, x + size, y - size, 48.0, CHART_LABEL_FONT);
            }
            prev = Some((x, y));
        }
    }
}

fn emit_pie<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_, '_>,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    let Some(series) = family.series.first() else {
        return;
    };
    let values: Vec<(usize, f64)> = (0..series.values.len().max(series.points.len()))
        .map(|index| (index, series_value(series, index)))
        .filter(|(_, value)| *value > 0.0 && value.is_finite())
        .take(ops.remaining)
        .collect();
    let total: f64 = values.iter().map(|(_, value)| value).sum();
    if total <= 0.0 {
        return;
    }
    let r = (width.min(height) * 0.34).max(10.0);
    let cx = x + width * 0.38;
    let cy = y + height * 0.46;
    let inner_r = if family.chart_type == "doughnut" {
        r * 0.48
    } else {
        0.0
    };
    let mut angle = -std::f64::consts::FRAC_PI_2;
    for (index, value) in &values {
        let sweep = (*value / total) * std::f64::consts::TAU;
        ops.push(PlotOp::Path {
            x: cx - r,
            y: cy - r,
            w: r * 2.0,
            h: r * 2.0,
            commands: pie_wedge_path(cx, cy, r, inner_r, angle, angle + sweep),
            fill: series_point_color(series, *index, *index),
            stroke: Some(PlotStroke {
                color: CHART_BACKGROUND_COLOR.to_owned(),
                width: 1.0,
            }),
        });
        angle += sweep;
    }
}

fn pie_wedge_path(
    cx: f64,
    cy: f64,
    r: f64,
    inner_r: f64,
    start: f64,
    end: f64,
) -> Vec<GeometryPathCommand> {
    let steps = (((end - start).abs() / std::f64::consts::TAU) * 48.0)
        .ceil()
        .max(2.0) as usize;
    let mut path = Vec::new();
    if inner_r > 0.0 {
        path.push(GeometryPathCommand::Move {
            x: cx + r * start.cos(),
            y: cy + r * start.sin(),
        });
    } else {
        path.push(GeometryPathCommand::Move { x: cx, y: cy });
        path.push(GeometryPathCommand::Line {
            x: cx + r * start.cos(),
            y: cy + r * start.sin(),
        });
    }
    for i in 1..=steps {
        let a = start + (end - start) * i as f64 / steps as f64;
        path.push(GeometryPathCommand::Line {
            x: cx + r * a.cos(),
            y: cy + r * a.sin(),
        });
    }
    if inner_r > 0.0 {
        path.push(GeometryPathCommand::Line {
            x: cx + inner_r * end.cos(),
            y: cy + inner_r * end.sin(),
        });
        for i in (0..steps).rev() {
            let a = start + (end - start) * i as f64 / steps as f64;
            path.push(GeometryPathCommand::Line {
                x: cx + inner_r * a.cos(),
                y: cy + inner_r * a.sin(),
            });
        }
    }
    path.push(GeometryPathCommand::Close);
    path
}

fn emit_legend<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    chart: &PlotChart<'_>,
    x: f64,
    y: f64,
    width: f64,
) {
    if !has_legend(chart) || width <= 0.0 {
        return;
    }
    let series: Vec<&PlotSeries<'_>> = if chart.series.is_empty() {
        chart
            .plot_groups
            .iter()
            .flat_map(|group| group.series.iter())
            .take(MAX_LEGEND_ENTRIES)
            .collect()
    } else {
        chart.series.iter().take(MAX_LEGEND_ENTRIES).collect()
    };
    let pie_legend = chart.chart_type == "pie"
        || chart.chart_type == "doughnut"
        || chart
            .plot_groups
            .iter()
            .any(|group| matches!(group.chart_type, Some("pie") | Some("doughnut")));
    let entries: Vec<(String, String)> = if pie_legend {
        series
            .as_slice()
            .first()
            .map(|series| {
                (0..series_length(series).min(MAX_LEGEND_ENTRIES))
                    .map(|i| {
                        (
                            series
                                .categories
                                .get(i)
                                .cloned()
                                .unwrap_or_else(|| (i + 1).to_string()),
                            series_point_color(series, i, i),
                        )
                    })
                    .collect()
            })
            .unwrap_or_default()
    } else {
        series
            .iter()
            .enumerate()
            .map(|(i, series)| {
                (
                    series
                        .name
                        .map(str::to_owned)
                        .unwrap_or_else(|| format!("Series {}", i + 1)),
                    series_color(Some(series), i),
                )
            })
            .collect()
    };
    for (i, (label, color)) in entries.iter().enumerate() {
        let yy = y + i as f64 * 15.0;
        push_rect(ops, x, yy, 8.0, 8.0, color);
        push_text(
            ops,
            label,
            x + 12.0,
            yy + 8.0,
            width - 12.0,
            CHART_LABEL_FONT,
        );
    }
}

fn format_number(value: f64) -> String {
    if value.abs() >= 100.0 || value.fract().abs() < 0.01 {
        format!("{value:.0}")
    } else {
        format!("{value:.1}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    struct Source {
        categories: Vec<String>,
        values: Vec<f64>,
    }

    fn source(values: &[f64]) -> Source {
        Source {
            categories: vec!["Q1".to_owned(), "Q2".to_owned()],
            values: values.to_vec(),
        }
    }

    fn series<'a>(name: &'a str, source: &'a Source) -> PlotSeries<'a> {
        PlotSeries {
            name: Some(name),
            categories: &source.categories,
            values: &source.values,
            ..PlotSeries::default()
        }
    }

    fn rect() -> PlotRect {
        PlotRect {
            x: 0.0,
            y: 0.0,
            w: 300.0,
            h: 200.0,
        }
    }

    fn family<'a, 'b>(chart: &'b PlotChart<'a>) -> PlotFamily<'a, 'b> {
        PlotFamily {
            chart_type: chart.chart_type,
            series: &chart.series,
            value_axis: chart.value_axis,
        }
    }

    #[test]
    fn fonts_render_the_css_shorthand() {
        assert_eq!(CHART_LABEL_FONT.css(), "400 10px Calibri, sans-serif");
        assert_eq!(CHART_TITLE_FONT.css(), "600 13px Calibri, sans-serif");
    }

    #[test]
    fn column_chart_emits_background_title_axes_bars_and_legend() {
        let north = source(&[10.0, 20.0]);
        let chart = PlotChart {
            chart_type: "column",
            title: Some("Revenue"),
            series: vec![series("North", &north)],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());

        assert!(matches!(&ops[0], PlotOp::Rect { fill, .. } if fill == CHART_BACKGROUND_COLOR));
        assert!(matches!(&ops[1], PlotOp::Text { text, font, .. }
            if text == "Revenue" && *font == CHART_TITLE_FONT));
        assert_eq!(
            ops.iter()
                .filter(|op| matches!(op, PlotOp::Line { .. }))
                .count(),
            7
        );
        assert!(
            ops.iter()
                .any(|op| matches!(op, PlotOp::Text { text, .. } if text == "North"))
        );
        assert_eq!(
            chart_aria_label(&chart),
            "Revenue, column chart, 1 series, 2 categories"
        );
    }

    #[test]
    fn pie_chart_emits_one_closed_wedge_per_positive_value() {
        let share = source(&[3.0, 0.0, 1.0]);
        let chart = PlotChart {
            chart_type: "pie",
            series: vec![series("Share", &share)],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        let paths: Vec<&PlotOp> = ops
            .iter()
            .filter(|op| matches!(op, PlotOp::Path { .. }))
            .collect();
        assert_eq!(paths.len(), 2);
        assert!(
            paths
                .iter()
                .all(|op| matches!(op, PlotOp::Path { commands, .. }
            if matches!(commands.last(), Some(GeometryPathCommand::Close))))
        );
    }

    #[test]
    fn axis_bounds_override_the_data_range() {
        let north = source(&[1.0]);
        let chart = PlotChart {
            chart_type: "column",
            value_axis: Some(PlotAxisRange {
                min: Some(-10.0),
                max: Some(10.0),
            }),
            series: vec![series("North", &north)],
            ..PlotChart::default()
        };
        assert_eq!(value_range(family(&chart)), (-10.0, 10.0));
    }

    #[test]
    fn point_marker_without_a_size_falls_back_to_the_default() {
        let north = source(&[1.0, 2.0]);
        let mut series = series("North", &north);
        series.marker = Some(PlotMarker { size: Some(9.0) });
        series.points = vec![PlotPoint {
            index: Some(0),
            marker: Some(PlotMarker { size: None }),
            ..PlotPoint::default()
        }];
        assert_eq!(marker_size(&series, 0), 4.0);
        assert_eq!(marker_size(&series, 1), 9.0);
    }

    #[test]
    fn combo_plot_groups_drive_the_label_and_the_legend() {
        let revenue = source(&[5.0, 9.0]);
        let trend = source(&[4.0, 8.0]);
        let chart = PlotChart {
            chart_type: "column",
            plot_groups: vec![
                PlotGroup {
                    chart_type: Some("column"),
                    series: vec![series("Revenue", &revenue)],
                    ..PlotGroup::default()
                },
                PlotGroup {
                    chart_type: Some("line"),
                    series: vec![series("Trend", &trend)],
                    ..PlotGroup::default()
                },
            ],
            ..PlotChart::default()
        };
        assert_eq!(
            chart_aria_label(&chart),
            "Untitled chart, combo chart, 2 series, 2 categories"
        );
        let ops = plot_chart(&chart, rect());
        assert!(
            ops.iter()
                .any(|op| matches!(op, PlotOp::Text { text, .. } if text == "Revenue"))
        );
        assert!(
            ops.iter()
                .any(|op| matches!(op, PlotOp::Text { text, .. } if text == "Trend"))
        );
    }

    #[test]
    fn the_op_budget_bounds_a_chart_with_far_more_data_than_it_can_draw() {
        let values: Vec<f64> = (0..200_000).map(|i| i as f64).collect();
        let wide = Source {
            categories: Vec::new(),
            values,
        };
        let chart = PlotChart {
            chart_type: "line",
            series: vec![series("Wide", &wide)],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        assert_eq!(ops.len(), MAX_PLOT_OPS);
    }

    #[test]
    fn the_sink_sees_the_same_ops_the_vector_wrapper_collects() {
        let north = source(&[10.0, 20.0]);
        let chart = PlotChart {
            chart_type: "column",
            title: Some("Revenue"),
            series: vec![series("North", &north)],
            ..PlotChart::default()
        };
        struct Counter(usize);
        impl PlotSink for Counter {
            fn push_op(&mut self, _: PlotOp) {
                self.0 += 1;
            }
        }
        let mut counter = Counter(0);
        plot_chart_into(&chart, rect(), &mut counter);
        assert_eq!(counter.0, plot_chart(&chart, rect()).len());
    }
}
