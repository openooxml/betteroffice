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
/// Hard ceiling on the data points one chart may index or scan for its range.
pub const MAX_PLOT_DATA_SCAN: usize = 200_000;
/// Hard ceiling on the plot groups one chart may draw.
pub const MAX_PLOT_GROUPS: usize = 64;
/// Hard ceiling on the series one chart may draw, across all its plot groups.
pub const MAX_PLOT_SERIES: usize = 1_024;
/// Coordinates are clamped here so every emitted op stays finite.
pub const MAX_PLOT_COORD: f64 = 1e9;

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
    pub axis_titles: PlotAxisTitles<'a>,
    pub series: Vec<PlotSeries<'a>>,
    pub plot_groups: Vec<PlotGroup<'a>>,
}

/// Axis titles, drawn horizontally because [`PlotOp::Text`] has no rotation.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotAxisTitles<'a> {
    pub category: Option<&'a str>,
    pub value: Option<&'a str>,
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
            axis_titles: PlotAxisTitles {
                category: space
                    .axes
                    .as_ref()
                    .and_then(|axes| axes.category.as_ref())
                    .and_then(|axis| axis.title.as_deref()),
                value: space
                    .axes
                    .as_ref()
                    .and_then(|axes| axes.value.as_ref())
                    .and_then(|axis| axis.title.as_deref()),
            },
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
            .filter_map(|point| {
                Some(PlotPoint {
                    index: match point.index {
                        Some(index) => Some(point_index(index)?),
                        None => None,
                    },
                    value: None,
                    color: Some(&point.color),
                    marker: None,
                    label: None,
                })
            })
            .collect(),
        grouping: series.grouping.as_deref(),
        marker: series
            .marker
            .as_ref()
            .map(|marker| PlotMarker { size: marker.size }),
    }
}

/// A `c:idx` that is a usable point index: an in-range non-negative integer.
/// Anything else is dropped rather than coerced, which would silently alias
/// point 0 or match every point.
fn point_index(index: f64) -> Option<usize> {
    (index.is_finite() && index >= 0.0 && index.fract() == 0.0 && index <= f64::from(u32::MAX))
        .then_some(index as usize)
}

/// Non-finite coordinates become zero and finite ones are clamped, so a
/// degenerate rectangle cannot produce NaN or infinite output.
fn finite(value: f64) -> f64 {
    if value.is_finite() {
        value.clamp(-MAX_PLOT_COORD, MAX_PLOT_COORD)
    } else {
        0.0
    }
}

/// Draw ops for `chart` inside `rect`, back to front, into `sink`.
pub fn plot_chart_into<S: PlotSink + ?Sized>(chart: &PlotChart<'_>, rect: PlotRect, sink: &mut S) {
    let (x, y, width, height) = (
        finite(rect.x),
        finite(rect.y),
        finite(rect.w),
        finite(rect.h),
    );
    let ops = &mut Emitter {
        sink,
        remaining: MAX_PLOT_OPS,
    };
    let scan = &mut ScanBudget::new();

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
        let series = series_views(&chart.series, scan);
        emit_family(
            ops,
            PlotFamily {
                chart_type: chart.chart_type,
                series: &series,
                value_axis: chart.value_axis,
                axis_titles: chart.axis_titles,
            },
            plot,
            x,
            y + title_h,
            width,
            height - title_h,
        );
    } else {
        for group in chart.plot_groups.iter().take(MAX_PLOT_GROUPS) {
            if ops.exhausted() {
                break;
            }
            let series = series_views(&group.series, scan);
            emit_family(
                ops,
                PlotFamily {
                    chart_type: group.chart_type.unwrap_or(chart.chart_type),
                    series: &series,
                    value_axis: chart.value_axis,
                    axis_titles: chart.axis_titles,
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
    emit_legend(
        ops,
        chart,
        scan,
        legend_x,
        y + title_h + 8.0,
        legend_w - 12.0,
    );
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
        chart.series.iter().map(series_length).max().unwrap_or(0)
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
struct PlotFamily<'a> {
    chart_type: &'a str,
    series: &'a [SeriesView<'a>],
    value_axis: Option<PlotAxisRange>,
    axis_titles: PlotAxisTitles<'a>,
}

/// A series plus a first-match index over its points, so a lookup costs a
/// binary search instead of a rescan.
struct SeriesView<'a> {
    series: &'a PlotSeries<'a>,
    /// Position of the first point without an index, which matches every query.
    wildcard: Option<usize>,
    /// `(point index, first position)`, sorted by point index.
    indexed: Vec<(usize, usize)>,
}

/// What one chart may still index, shared across its plot groups so a
/// per-series limit cannot multiply by the series count.
struct ScanBudget {
    series: usize,
    points: usize,
}

impl ScanBudget {
    fn new() -> Self {
        Self {
            series: MAX_PLOT_SERIES,
            points: MAX_PLOT_DATA_SCAN,
        }
    }

    fn take_series(&mut self, requested: usize) -> usize {
        let allowed = requested.min(self.series);
        self.series -= allowed;
        allowed
    }

    fn take_points(&mut self, requested: usize) -> usize {
        let allowed = requested.min(self.points);
        self.points -= allowed;
        allowed
    }
}

fn series_views<'a>(series: &'a [PlotSeries<'a>], budget: &mut ScanBudget) -> Vec<SeriesView<'a>> {
    let allowed = budget.take_series(series.len());
    series
        .iter()
        .take(allowed)
        .map(|series| SeriesView::new(series, budget))
        .collect()
}

impl<'a> SeriesView<'a> {
    fn new(series: &'a PlotSeries<'a>, budget: &mut ScanBudget) -> Self {
        let mut wildcard = None;
        let scanned = budget.take_points(series.points.len());
        let mut indexed = Vec::with_capacity(scanned);
        for (position, point) in series.points.iter().take(scanned).enumerate() {
            match point.index {
                Some(index) => indexed.push((index, position)),
                None if wildcard.is_none() => wildcard = Some(position),
                None => {}
            }
        }
        indexed.sort_unstable();
        indexed.dedup_by_key(|(index, _)| *index);
        Self {
            series,
            wildcard,
            indexed,
        }
    }

    fn point(&self, index: usize) -> Option<&PlotPoint<'a>> {
        let exact = self
            .indexed
            .binary_search_by_key(&index, |(key, _)| *key)
            .ok()
            .map(|slot| self.indexed[slot].1);
        let position = match (self.wildcard, exact) {
            (Some(wildcard), Some(exact)) => Some(wildcard.min(exact)),
            (wildcard, exact) => wildcard.or(exact),
        };
        position.map(|position| &self.series.points[position])
    }

    fn value(&self, index: usize) -> f64 {
        self.point(index)
            .and_then(|point| point.value)
            .or_else(|| self.series.values.get(index).copied())
            .filter(|value| value.is_finite())
            .unwrap_or(0.0)
    }

    fn point_color(&self, point_index: usize, series_index: usize) -> String {
        self.point(point_index)
            .and_then(|point| point.color)
            .map(hex)
            .unwrap_or_else(|| series_color(Some(self.series), series_index))
    }

    fn marker_size(&self, index: usize) -> f64 {
        self.point(index)
            .and_then(|point| point.marker.as_ref())
            .or(self.series.marker.as_ref())
            .and_then(|marker| marker.size)
            .unwrap_or(4.0)
            .clamp(1.0, 24.0)
    }

    fn length(&self) -> usize {
        series_length(self.series)
    }
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
    family: PlotFamily<'_>,
    plot: PlotArea,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    match family.chart_type {
        "pie" | "doughnut" | "ofPie" => emit_pie(ops, family, x, y, width, height),
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

fn series_length(series: &PlotSeries<'_>) -> usize {
    series
        .categories
        .len()
        .max(series.values.len())
        .max(series.points.len())
}

fn category_count(series: &[SeriesView<'_>]) -> usize {
    series.iter().map(SeriesView::length).max().unwrap_or(0)
}

fn category_label(series: &[SeriesView<'_>], index: usize) -> String {
    series
        .iter()
        .find_map(|series| series.series.categories.get(index).cloned())
        .unwrap_or_else(|| (index + 1).to_string())
}

fn value_range(family: PlotFamily<'_>) -> (f64, f64) {
    let mut min = 0.0;
    let mut max = 0.0;
    let mut remaining = MAX_PLOT_DATA_SCAN;
    for series in family.series {
        let samples = series
            .series
            .values
            .len()
            .max(series.series.points.len())
            .min(remaining);
        remaining -= samples;
        for index in 0..samples {
            let value = series.value(index);
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
    if !(max - min).is_finite() || max <= min {
        (min, max) = (0.0, 1.0);
    }
    (min, max)
}

fn value_y(plot: PlotArea, value: f64, min: f64, max: f64) -> f64 {
    plot.y + (max - value) / (max - min) * plot.h
}

fn emit_axes<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
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
    if let Some(title) = family.axis_titles.value.filter(|title| !title.is_empty()) {
        push_text(
            ops,
            title,
            plot.x - 38.0,
            plot.y - 5.0,
            plot.w + 38.0,
            CHART_LABEL_FONT,
        );
    }
    if let Some(title) = family
        .axis_titles
        .category
        .filter(|title| !title.is_empty())
    {
        push_text(
            ops,
            title,
            plot.x,
            plot.y + plot.h + 26.0,
            plot.w,
            CHART_LABEL_FONT,
        );
    }
}

fn emit_bar<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
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
                let value = series.value(cat_idx);
                let ratio = ((value - min) / (max - min)).clamp(0.0, 1.0);
                let bar_w = ratio * plot.w;
                let y = plot.y + row_h * cat_idx as f64 + row_h * 0.15 + bar_h * ser_idx as f64;
                push_rect(
                    ops,
                    plot.x,
                    y,
                    bar_w,
                    bar_h,
                    &series.point_color(cat_idx, ser_idx),
                );
                if let Some(label) = series.point(cat_idx).and_then(|point| point.label) {
                    push_text(
                        ops,
                        label,
                        plot.x + bar_w + 3.0,
                        y + bar_h,
                        48.0,
                        CHART_LABEL_FONT,
                    );
                }
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
                let value = series.value(cat_idx);
                let yv = value_y(plot, value.clamp(min, max), min, max);
                let y0 = zero_y;
                let x = plot.x + group_w * cat_idx as f64 + group_w * 0.15 + bar_w * ser_idx as f64;
                push_rect(
                    ops,
                    x,
                    yv.min(y0),
                    bar_w,
                    (y0 - yv).abs().max(1.0),
                    &series.point_color(cat_idx, ser_idx),
                );
                if let Some(label) = series.point(cat_idx).and_then(|point| point.label) {
                    push_text(
                        ops,
                        label,
                        x,
                        yv.min(y0) - 3.0,
                        bar_w.max(32.0),
                        CHART_LABEL_FONT,
                    );
                }
            }
        }
    }
}

fn emit_line<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
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
        let color = series_color(Some(series.series), ser_idx);
        let mut prev: Option<(f64, f64)> = None;
        for i in 0..cat_count {
            if ops.exhausted() {
                return;
            }
            let value = series.value(i);
            let x = plot.x + plot.w * i as f64 / denom;
            let y = value_y(plot, value.clamp(min, max), min, max);
            if let Some((prev_x, prev_y)) = prev {
                push_line(ops, prev_x, prev_y, x, y, &color, 2.0);
            }
            let size = series.marker_size(i);
            let point_color = series.point_color(i, ser_idx);
            push_rect(
                ops,
                x - size / 2.0,
                y - size / 2.0,
                size,
                size,
                &point_color,
            );
            if let Some(label) = series.point(i).and_then(|point| point.label) {
                push_text(ops, label, x + size, y - size, 48.0, CHART_LABEL_FONT);
            }
            prev = Some((x, y));
        }
    }
}

fn emit_pie<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    let Some(series) = family.series.first() else {
        return;
    };
    let scanned = series
        .series
        .values
        .len()
        .max(series.series.points.len())
        .min(MAX_PLOT_DATA_SCAN);
    let values: Vec<(usize, f64)> = (0..scanned)
        .map(|index| (index, series.value(index)))
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
            fill: series.point_color(*index, *index),
            stroke: Some(PlotStroke {
                color: CHART_BACKGROUND_COLOR.to_owned(),
                width: 1.0,
            }),
        });
        if let Some(label) = series.point(*index).and_then(|point| point.label) {
            let middle = angle + sweep / 2.0;
            push_text(
                ops,
                label,
                cx + r * 0.62 * middle.cos(),
                cy + r * 0.62 * middle.sin(),
                48.0,
                CHART_LABEL_FONT,
            );
        }
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
    budget: &mut ScanBudget,
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
    let pie_legend = matches!(chart.chart_type, "pie" | "doughnut" | "ofPie")
        || chart.plot_groups.iter().any(|group| {
            matches!(
                group.chart_type,
                Some("pie") | Some("doughnut") | Some("ofPie")
            )
        });
    let entries: Vec<(String, String)> = if pie_legend {
        series
            .as_slice()
            .first()
            .map(|series| SeriesView::new(series, budget))
            .map(|series| {
                (0..series.length().min(MAX_LEGEND_ENTRIES))
                    .map(|i| {
                        (
                            series
                                .series
                                .categories
                                .get(i)
                                .cloned()
                                .unwrap_or_else(|| (i + 1).to_string()),
                            series.point_color(i, i),
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

/// Axis-tick formatting, shared with hosts that inject value data labels.
pub fn format_number(value: f64) -> String {
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

    fn family<'a>(chart: &'a PlotChart<'a>, series: &'a [SeriesView<'a>]) -> PlotFamily<'a> {
        PlotFamily {
            chart_type: chart.chart_type,
            series,
            value_axis: chart.value_axis,
            axis_titles: chart.axis_titles,
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
    fn axis_titles_draw_beside_the_axes_they_name() {
        let north = source(&[10.0, 20.0]);
        for chart_type in ["column", "bar", "line"] {
            let chart = PlotChart {
                chart_type,
                axis_titles: PlotAxisTitles {
                    category: Some("Quarter"),
                    value: Some("Millions"),
                },
                series: vec![series("North", &north)],
                ..PlotChart::default()
            };
            let ops = plot_chart(&chart, rect());
            for title in ["Quarter", "Millions"] {
                assert!(
                    ops.iter()
                        .any(|op| matches!(op, PlotOp::Text { text, .. } if text == title)),
                    "{chart_type} drops {title}"
                );
            }
        }
        let pie = PlotChart {
            chart_type: "pie",
            axis_titles: PlotAxisTitles {
                category: Some("Quarter"),
                value: Some("Millions"),
            },
            series: vec![series("North", &north)],
            ..PlotChart::default()
        };
        assert!(
            !plot_chart(&pie, rect())
                .iter()
                .any(|op| matches!(op, PlotOp::Text { text, .. } if text == "Millions"))
        );
    }

    #[test]
    fn an_of_pie_group_draws_wedges_and_a_per_slice_legend() {
        let share = source(&[3.0, 1.0]);
        let chart = PlotChart {
            chart_type: "pie",
            plot_groups: vec![PlotGroup {
                chart_type: Some("ofPie"),
                series: vec![series("Share", &share)],
                ..PlotGroup::default()
            }],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        assert_eq!(
            ops.iter()
                .filter(|op| matches!(op, PlotOp::Path { .. }))
                .count(),
            2
        );
        assert!(
            ops.iter()
                .any(|op| matches!(op, PlotOp::Text { text, .. } if text == "Q1"))
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
        let views = series_views(&chart.series, &mut ScanBudget::new());
        assert_eq!(value_range(family(&chart, &views)), (-10.0, 10.0));
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
        let view = SeriesView::new(&series, &mut ScanBudget::new());
        assert_eq!(view.marker_size(0), 4.0);
        assert_eq!(view.marker_size(1), 9.0);
    }

    #[test]
    fn point_labels_draw_on_every_chart_family() {
        let data = source(&[3.0, 1.0]);
        for chart_type in ["column", "bar", "line", "pie", "doughnut"] {
            let mut labelled = series("North", &data);
            labelled.points = vec![
                PlotPoint {
                    index: Some(0),
                    label: Some("first"),
                    ..PlotPoint::default()
                },
                PlotPoint {
                    index: Some(1),
                    label: Some("second"),
                    ..PlotPoint::default()
                },
            ];
            let chart = PlotChart {
                chart_type,
                series: vec![labelled],
                ..PlotChart::default()
            };
            let ops = plot_chart(&chart, rect());
            for label in ["first", "second"] {
                assert!(
                    ops.iter()
                        .any(|op| matches!(op, PlotOp::Text { text, .. } if text == label)),
                    "{chart_type} drops {label}"
                );
            }
        }
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
    fn the_index_budget_is_chart_wide_rather_than_per_series() {
        let data = source(&[1.0, 2.0]);
        let mut indexed = series("Indexed", &data);
        indexed.points = vec![PlotPoint {
            index: Some(1),
            color: Some("#010203"),
            ..PlotPoint::default()
        }];
        let both = [indexed.clone(), indexed];

        let views = series_views(&both, &mut ScanBudget::new());
        assert_eq!(views[1].indexed.len(), 1);

        let mut spent = ScanBudget::new();
        assert_eq!(spent.take_points(MAX_PLOT_DATA_SCAN), MAX_PLOT_DATA_SCAN);
        assert_eq!(spent.take_points(1), 0);
        assert!(series_views(&both, &mut spent)[1].indexed.is_empty());

        let mut capped = ScanBudget::new();
        assert_eq!(capped.take_series(MAX_PLOT_SERIES + 1), MAX_PLOT_SERIES);
        assert!(series_views(&both, &mut capped).is_empty());
    }

    #[test]
    fn the_series_and_group_caps_bound_a_directly_constructed_chart() {
        let data = source(&[1.0, 2.0]);
        let drawn = PlotSeries {
            color: Some("#111111"),
            ..series("Drawn", &data)
        };
        let beyond = PlotSeries {
            color: Some("#ABCDEF"),
            ..series("Beyond", &data)
        };
        fn group<'a>(series: Vec<PlotSeries<'a>>) -> PlotGroup<'a> {
            PlotGroup {
                chart_type: Some("column"),
                series,
                ..PlotGroup::default()
            }
        }

        let mut wide = vec![drawn.clone(); MAX_PLOT_SERIES];
        wide.push(beyond.clone());
        let mut many = vec![group(vec![drawn]); MAX_PLOT_GROUPS];
        many.push(group(vec![beyond]));

        for plot_groups in [vec![group(wide)], many] {
            let chart = PlotChart {
                chart_type: "column",
                plot_groups,
                ..PlotChart::default()
            };
            let ops = plot_chart(&chart, rect());
            assert!(ops.len() < MAX_PLOT_OPS);
            assert!(
                ops.iter()
                    .all(|op| !matches!(op, PlotOp::Rect { fill, .. } if fill == "#ABCDEF"))
            );
        }
    }

    #[test]
    fn a_later_plot_group_adds_nothing_once_the_op_budget_is_spent() {
        let values: Vec<f64> = (0..200_000).map(|i| i as f64).collect();
        let wide = Source {
            categories: Vec::new(),
            values,
        };
        let tail = source(&[1.0, 2.0]);
        let chart = PlotChart {
            chart_type: "line",
            plot_groups: vec![
                PlotGroup {
                    chart_type: Some("line"),
                    series: vec![series("Wide", &wide)],
                    ..PlotGroup::default()
                },
                PlotGroup {
                    chart_type: Some("column"),
                    series: vec![PlotSeries {
                        color: Some("#ABCDEF"),
                        ..series("Tail", &tail)
                    }],
                    ..PlotGroup::default()
                },
            ],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        assert_eq!(ops.len(), MAX_PLOT_OPS);
        assert!(
            ops.iter()
                .all(|op| !matches!(op, PlotOp::Rect { fill, .. } if fill == "#ABCDEF"))
        );
    }

    #[test]
    fn the_point_index_agrees_with_a_linear_first_match_scan() {
        let data = source(&[1.0, 2.0, 3.0, 4.0]);
        let cases: [Vec<Option<usize>>; 7] = [
            vec![],
            vec![Some(0), Some(1), Some(2)],
            vec![Some(2), Some(0)],
            vec![None, Some(0)],
            vec![Some(1), None, Some(0)],
            vec![Some(3), Some(3), None, None],
            vec![Some(9), Some(1), Some(9)],
        ];
        for indexes in cases {
            let mut series = series("S", &data);
            series.points = indexes
                .iter()
                .enumerate()
                .map(|(position, index)| PlotPoint {
                    index: *index,
                    value: Some(position as f64),
                    ..PlotPoint::default()
                })
                .collect();
            let view = SeriesView::new(&series, &mut ScanBudget::new());
            for query in 0..12 {
                let expected = series
                    .points
                    .iter()
                    .find(|point| point.index.unwrap_or(query) == query);
                assert_eq!(view.point(query), expected, "{indexes:?} at {query}");
            }
        }
    }

    #[test]
    fn a_point_per_category_resolves_without_rescanning_the_point_vector() {
        let values: Vec<f64> = (0..20_000).map(|i| (i % 13) as f64).collect();
        let dense = Source {
            categories: Vec::new(),
            values,
        };
        let mut series = series("Dense", &dense);
        series.points = (0..20_000)
            .map(|index| PlotPoint {
                index: Some(index),
                color: Some("#010203"),
                ..PlotPoint::default()
            })
            .collect();
        let chart = PlotChart {
            chart_type: "line",
            series: vec![series],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        assert!(ops.len() <= MAX_PLOT_OPS);
        assert!(
            ops.iter()
                .any(|op| matches!(op, PlotOp::Rect { fill, .. } if fill == "#010203"))
        );
    }

    #[test]
    fn out_of_schema_point_indexes_are_dropped_rather_than_coerced() {
        let space = ChartSpace {
            chart_type: "pie".to_owned(),
            title: None,
            legend: None,
            series: Vec::new(),
            axes: None,
            axis_list: None,
            plot_groups: vec![crate::chart::ChartPlotGroup {
                chart_type: Some("pie".to_owned()),
                grouping: None,
                overlap: None,
                gap_width: None,
                axis_ids: Vec::new(),
                vary_colors: false,
                first_slice_angle: None,
                hole_size: None,
                data_labels: None,
                series: vec![crate::chart::ChartSeries {
                    name: None,
                    categories: Vec::new(),
                    values: vec![1.0, 2.0],
                    color: "#4472C4".to_owned(),
                    index: None,
                    order: None,
                    category_formula: None,
                    value_formula: None,
                    axis_ids: None,
                    grouping: None,
                    marker: None,
                    smooth: None,
                    data_labels: None,
                    points: Some(
                        [-1.0, 1.5, 1e30, f64::NAN, 2.0]
                            .into_iter()
                            .map(|index| crate::chart::ChartPoint {
                                index: Some(index),
                                explosion: None,
                                color: "#010203".to_owned(),
                            })
                            .collect(),
                    ),
                }],
            }],
        };
        let chart = PlotChart::from(&space);
        let points = &chart.plot_groups[0].series[0].points;
        assert_eq!(points.len(), 1);
        assert_eq!(points[0].index, Some(2));
    }

    #[test]
    fn degenerate_rects_and_extreme_ranges_stay_finite() {
        let extremes = Source {
            categories: Vec::new(),
            values: vec![f64::MAX, -f64::MAX, 5.0],
        };
        let rects = [
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
                h: 200.0,
            },
            PlotRect {
                x: 0.0,
                y: 0.0,
                w: 1e308,
                h: 1e308,
            },
            PlotRect::default(),
        ];
        for chart_type in ["column", "bar", "line", "pie"] {
            let chart = PlotChart {
                chart_type,
                series: vec![series("Extreme", &extremes)],
                ..PlotChart::default()
            };
            for rect in rects {
                for op in plot_chart(&chart, rect) {
                    assert!(op_is_finite(&op), "{chart_type} {rect:?} {op:?}");
                }
            }
        }
    }

    fn op_is_finite(op: &PlotOp) -> bool {
        let numbers: Vec<f64> = match op {
            PlotOp::Rect { x, y, w, h, .. } => vec![*x, *y, *w, *h],
            PlotOp::Text {
                x,
                baseline_y,
                width,
                ..
            } => vec![*x, *baseline_y, *width],
            PlotOp::Line { x1, y1, x2, y2, .. } => vec![*x1, *y1, *x2, *y2],
            PlotOp::Path {
                x,
                y,
                w,
                h,
                commands,
                ..
            } => {
                let mut numbers = vec![*x, *y, *w, *h];
                for command in commands {
                    match command {
                        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                            numbers.extend([*x, *y])
                        }
                        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                            numbers.extend([*cpx, *cpy, *x, *y]);
                        }
                        GeometryPathCommand::Cubic {
                            cp1x,
                            cp1y,
                            cp2x,
                            cp2y,
                            x,
                            y,
                        } => numbers.extend([*cp1x, *cp1y, *cp2x, *cp2y, *x, *y]),
                        GeometryPathCommand::Close => {}
                    }
                }
                numbers
            }
        };
        numbers.iter().all(|number| number.is_finite())
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
