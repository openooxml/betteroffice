//! Format-neutral chart geometry: a chart plus a rectangle in, an ordered
//! [`PlotOp`] list out. Hosts translate the ops into their own primitives.

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

/// What the geometry reads off a chart. Distinct from [`ChartSpace`], which is
/// the parse-fidelity model: hosts may inject per-point values and labels that
/// never appear in the chart part.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotChart {
    pub chart_type: String,
    pub title: Option<String>,
    pub legend: Option<PlotLegend>,
    pub value_axis: Option<PlotAxisRange>,
    pub series: Vec<PlotSeries>,
    pub plot_groups: Vec<PlotGroup>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotLegend {
    pub position: Option<String>,
    pub visible: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotAxisRange {
    pub min: Option<f64>,
    pub max: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotGroup {
    pub chart_type: Option<String>,
    pub grouping: Option<String>,
    pub series: Vec<PlotSeries>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotSeries {
    pub name: Option<String>,
    pub categories: Vec<String>,
    pub values: Vec<f64>,
    pub color: Option<String>,
    pub points: Vec<PlotPoint>,
    pub grouping: Option<String>,
    pub marker: Option<PlotMarker>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotPoint {
    pub index: Option<usize>,
    pub value: Option<f64>,
    pub color: Option<String>,
    pub marker: Option<PlotMarker>,
    pub label: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotMarker {
    pub size: Option<f64>,
}

impl From<&ChartSpace> for PlotChart {
    fn from(space: &ChartSpace) -> Self {
        Self {
            chart_type: space.chart_type.clone(),
            title: space.title.clone(),
            legend: space.legend.as_ref().map(|legend| PlotLegend {
                position: legend.position.clone(),
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
                    chart_type: group.chart_type.clone(),
                    grouping: group.grouping.clone(),
                    series: group.series.iter().map(plot_series_from_model).collect(),
                })
                .collect(),
        }
    }
}

fn plot_series_from_model(series: &super::model::ChartSeries) -> PlotSeries {
    PlotSeries {
        name: series.name.clone(),
        categories: series.categories.clone(),
        values: series.values.clone(),
        color: Some(series.color.clone()),
        points: series
            .points
            .iter()
            .flatten()
            .map(|point| PlotPoint {
                index: point.index.map(|index| index as usize),
                value: None,
                color: Some(point.color.clone()),
                marker: None,
                label: None,
            })
            .collect(),
        grouping: series.grouping.clone(),
        marker: series
            .marker
            .as_ref()
            .map(|marker| PlotMarker { size: marker.size }),
    }
}

/// Draw ops for `chart` inside `rect`, back to front.
pub fn plot_chart(chart: &PlotChart, rect: PlotRect) -> Vec<PlotOp> {
    let PlotRect {
        x,
        y,
        w: width,
        h: height,
    } = rect;
    let mut ops = Vec::new();

    push_rect(&mut ops, x, y, width, height, CHART_BACKGROUND_COLOR);

    let title_h = if let Some(title) = chart.title.as_deref().filter(|s| !s.is_empty()) {
        push_text(
            &mut ops,
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
        .and_then(|legend| legend.position.as_deref())
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
            &mut ops,
            chart,
            plot,
            x,
            y + title_h,
            width,
            height - title_h,
        );
    } else {
        for group in &chart.plot_groups {
            let mut group_chart = chart.clone();
            group_chart.chart_type = group
                .chart_type
                .clone()
                .unwrap_or_else(|| chart.chart_type.clone());
            group_chart.series = group.series.clone();
            for series in &mut group_chart.series {
                if series.grouping.is_none() {
                    series.grouping = group.grouping.clone();
                }
            }
            group_chart.plot_groups.clear();
            emit_family(
                &mut ops,
                &group_chart,
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
        &mut ops,
        chart,
        legend_x,
        y + title_h + 8.0,
        legend_w - 12.0,
    );
    ops
}

/// Screen-reader summary of a chart.
pub fn chart_aria_label(chart: &PlotChart) -> String {
    let kind = if chart.plot_groups.len() > 1 {
        "combo chart"
    } else {
        match chart.chart_type.as_str() {
            "bar" => "bar chart",
            "line" => "line chart",
            "pie" => "pie chart",
            "doughnut" => "doughnut chart",
            _ => "column chart",
        }
    };
    let title = chart
        .title
        .as_deref()
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
        category_count(chart)
    };
    format!("{title}, {kind}, {series_count} series, {category_count} categories")
}

#[derive(Clone, Copy)]
struct PlotArea {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}

fn emit_family(
    ops: &mut Vec<PlotOp>,
    chart: &PlotChart,
    plot: PlotArea,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    match chart.chart_type.as_str() {
        "pie" | "doughnut" => emit_pie(ops, chart, x, y, width, height),
        "line" | "scatter" | "radar" => emit_line(ops, chart, plot),
        "bar" => emit_bar(ops, chart, plot, true),
        _ => emit_bar(ops, chart, plot, false),
    }
}

fn push_rect(ops: &mut Vec<PlotOp>, x: f64, y: f64, w: f64, h: f64, fill: &str) {
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

fn push_text(
    ops: &mut Vec<PlotOp>,
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

fn push_line(ops: &mut Vec<PlotOp>, x1: f64, y1: f64, x2: f64, y2: f64, color: &str, width: f64) {
    ops.push(PlotOp::Line {
        x1,
        y1,
        x2,
        y2,
        color: color.to_owned(),
        width,
    });
}

fn has_legend(chart: &PlotChart) -> bool {
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

fn series_color(series: Option<&PlotSeries>, index: usize) -> String {
    series
        .and_then(|series| series.color.as_deref())
        .filter(|color| !color.is_empty())
        .map(hex)
        .unwrap_or_else(|| CHART_SERIES_COLORS[index % CHART_SERIES_COLORS.len()].to_owned())
}

fn series_point(series: &PlotSeries, index: usize) -> Option<&PlotPoint> {
    series
        .points
        .iter()
        .find(|point| point.index.unwrap_or(index) == index)
}

fn series_value(series: &PlotSeries, index: usize) -> f64 {
    series_point(series, index)
        .and_then(|point| point.value)
        .or_else(|| series.values.get(index).copied())
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
}

fn series_point_color(series: &PlotSeries, point_index: usize, series_index: usize) -> String {
    series_point(series, point_index)
        .and_then(|point| point.color.as_deref())
        .map(hex)
        .unwrap_or_else(|| series_color(Some(series), series_index))
}

fn series_length(series: &PlotSeries) -> usize {
    series
        .categories
        .len()
        .max(series.values.len())
        .max(series.points.len())
}

fn category_count(chart: &PlotChart) -> usize {
    chart.series.iter().map(series_length).max().unwrap_or(0)
}

fn category_label(chart: &PlotChart, index: usize) -> String {
    chart
        .series
        .iter()
        .find_map(|series| series.categories.get(index).cloned())
        .unwrap_or_else(|| (index + 1).to_string())
}

fn value_range(chart: &PlotChart) -> (f64, f64) {
    let mut min = 0.0;
    let mut max = 0.0;
    for series in &chart.series {
        for index in 0..series.values.len().max(series.points.len()) {
            let value = series_value(series, index);
            if value.is_finite() {
                min = f64::min(min, value);
                max = f64::max(max, value);
            }
        }
    }
    if let Some(axis) = chart.value_axis.as_ref() {
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

fn marker_size(series: &PlotSeries, index: usize) -> f64 {
    series_point(series, index)
        .and_then(|point| point.marker.as_ref())
        .or(series.marker.as_ref())
        .and_then(|marker| marker.size)
        .unwrap_or(4.0)
        .clamp(1.0, 24.0)
}

fn emit_axes(ops: &mut Vec<PlotOp>, chart: &PlotChart, plot: PlotArea) {
    let (min, max) = value_range(chart);
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

fn emit_bar(ops: &mut Vec<PlotOp>, chart: &PlotChart, plot: PlotArea, horizontal: bool) {
    let cat_count = category_count(chart);
    if cat_count == 0 || chart.series.is_empty() {
        return;
    }
    emit_axes(ops, chart, plot);
    let (min, max) = value_range(chart);
    let zero_y = value_y(plot, 0.0_f64.clamp(min, max), min, max);
    let series_count = chart.series.len().max(1);
    if horizontal {
        let row_h = plot.h / cat_count as f64;
        let bar_h = (row_h * 0.7 / series_count as f64).max(1.0);
        for cat_idx in 0..cat_count {
            let label = category_label(chart, cat_idx);
            push_text(
                ops,
                &label,
                plot.x - 38.0,
                plot.y + row_h * (cat_idx as f64 + 0.55),
                36.0,
                CHART_LABEL_FONT,
            );
            for (ser_idx, series) in chart.series.iter().enumerate() {
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
            let label = category_label(chart, cat_idx);
            push_text(
                ops,
                &label,
                plot.x + group_w * cat_idx as f64 + 2.0,
                plot.y + plot.h + 14.0,
                group_w - 4.0,
                CHART_LABEL_FONT,
            );
            for (ser_idx, series) in chart.series.iter().enumerate() {
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

fn emit_line(ops: &mut Vec<PlotOp>, chart: &PlotChart, plot: PlotArea) {
    let cat_count = category_count(chart);
    if cat_count == 0 || chart.series.is_empty() {
        return;
    }
    emit_axes(ops, chart, plot);
    let (min, max) = value_range(chart);
    let denom = (cat_count.saturating_sub(1)).max(1) as f64;
    for i in 0..cat_count {
        let label = category_label(chart, i);
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
    for (ser_idx, series) in chart.series.iter().enumerate() {
        let color = series_color(Some(series), ser_idx);
        let mut prev: Option<(f64, f64)> = None;
        for i in 0..cat_count {
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
            if let Some(label) = series_point(series, i).and_then(|point| point.label.as_deref()) {
                push_text(ops, label, x + size, y - size, 48.0, CHART_LABEL_FONT);
            }
            prev = Some((x, y));
        }
    }
}

fn emit_pie(ops: &mut Vec<PlotOp>, chart: &PlotChart, x: f64, y: f64, width: f64, height: f64) {
    let Some(series) = chart.series.first() else {
        return;
    };
    let values: Vec<(usize, f64)> = (0..series.values.len().max(series.points.len()))
        .map(|index| (index, series_value(series, index)))
        .filter(|(_, value)| *value > 0.0 && value.is_finite())
        .collect();
    let total: f64 = values.iter().map(|(_, value)| value).sum();
    if total <= 0.0 {
        return;
    }
    let r = (width.min(height) * 0.34).max(10.0);
    let cx = x + width * 0.38;
    let cy = y + height * 0.46;
    let inner_r = if chart.chart_type == "doughnut" {
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

fn emit_legend(ops: &mut Vec<PlotOp>, chart: &PlotChart, x: f64, y: f64, width: f64) {
    if !has_legend(chart) || width <= 0.0 {
        return;
    }
    let series: Vec<&PlotSeries> = if chart.series.is_empty() {
        chart
            .plot_groups
            .iter()
            .flat_map(|group| group.series.iter())
            .collect()
    } else {
        chart.series.iter().collect()
    };
    let pie_legend = chart.chart_type == "pie"
        || chart.chart_type == "doughnut"
        || chart
            .plot_groups
            .iter()
            .any(|group| matches!(group.chart_type.as_deref(), Some("pie") | Some("doughnut")));
    let entries: Vec<(String, String)> = if pie_legend {
        series
            .as_slice()
            .first()
            .map(|series| {
                (0..series_length(series))
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
                        .clone()
                        .unwrap_or_else(|| format!("Series {}", i + 1)),
                    series_color(Some(series), i),
                )
            })
            .collect()
    };
    for (i, (label, color)) in entries.iter().take(MAX_LEGEND_ENTRIES).enumerate() {
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

    fn series(name: &str, values: Vec<f64>) -> PlotSeries {
        PlotSeries {
            name: Some(name.to_owned()),
            categories: vec!["Q1".to_owned(), "Q2".to_owned()],
            values,
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

    #[test]
    fn fonts_render_the_css_shorthand() {
        assert_eq!(CHART_LABEL_FONT.css(), "400 10px Calibri, sans-serif");
        assert_eq!(CHART_TITLE_FONT.css(), "600 13px Calibri, sans-serif");
    }

    #[test]
    fn column_chart_emits_background_title_axes_bars_and_legend() {
        let chart = PlotChart {
            chart_type: "column".to_owned(),
            title: Some("Revenue".to_owned()),
            series: vec![series("North", vec![10.0, 20.0])],
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
        let chart = PlotChart {
            chart_type: "pie".to_owned(),
            series: vec![series("Share", vec![3.0, 0.0, 1.0])],
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
        let chart = PlotChart {
            chart_type: "column".to_owned(),
            value_axis: Some(PlotAxisRange {
                min: Some(-10.0),
                max: Some(10.0),
            }),
            series: vec![series("North", vec![1.0])],
            ..PlotChart::default()
        };
        assert_eq!(value_range(&chart), (-10.0, 10.0));
    }

    #[test]
    fn point_marker_without_a_size_falls_back_to_the_default() {
        let mut series = series("North", vec![1.0, 2.0]);
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
        let chart = PlotChart {
            chart_type: "column".to_owned(),
            plot_groups: vec![
                PlotGroup {
                    chart_type: Some("column".to_owned()),
                    series: vec![series("Revenue", vec![5.0, 9.0])],
                    ..PlotGroup::default()
                },
                PlotGroup {
                    chart_type: Some("line".to_owned()),
                    series: vec![series("Trend", vec![4.0, 8.0])],
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
}
