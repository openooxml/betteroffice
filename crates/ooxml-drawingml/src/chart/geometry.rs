//! Format-neutral chart geometry: a chart plus a rectangle in, an ordered
//! [`PlotOp`] stream out. Hosts translate the ops into their own primitives.

use std::collections::{HashMap, HashSet};

use crate::GeometryPathCommand;

use super::model::ChartSpace;

pub const CHART_AXIS_COLOR: &str = "#666666";
pub const CHART_GRID_COLOR: &str = "#D9D9D9";
pub const CHART_TEXT_COLOR: &str = "#222222";
pub const CHART_BACKGROUND_COLOR: &str = "#FFFFFF";
pub const CHART_SERIES_COLORS: [&str; 8] = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E",
];
/// The font family every chart text falls back to when `c:txPr` names none.
pub const CHART_FONT_FAMILY: &str = "Calibri, sans-serif";
pub const CHART_LABEL_SIZE_PX: f64 = 10.0;
pub const CHART_TITLE_SIZE_PX: f64 = 13.0;

/// The label font of a chart whose `c:txPr` says nothing.
pub fn chart_label_font() -> PlotFont {
    PlotFont {
        weight: 400,
        size_px: CHART_LABEL_SIZE_PX,
        family: CHART_FONT_FAMILY.to_owned(),
        italic: false,
    }
}

/// The title font of a chart whose `c:txPr` says nothing.
pub fn chart_title_font() -> PlotFont {
    PlotFont {
        weight: 600,
        size_px: CHART_TITLE_SIZE_PX,
        family: CHART_FONT_FAMILY.to_owned(),
        italic: false,
    }
}

/// Hard ceiling on the ops one chart may emit, whatever its data length.
pub const MAX_PLOT_OPS: usize = 100_000;
/// Hard ceiling on the data points one chart may index or scan for its range.
pub const MAX_PLOT_DATA_SCAN: usize = 200_000;
/// Hard ceiling on the plot groups one chart may draw.
pub const MAX_PLOT_GROUPS: usize = 64;
/// Hard ceiling on the series one chart may draw, across all its plot groups.
pub const MAX_PLOT_SERIES: usize = 1_024;
/// Hard ceiling on the ticks one axis may place, whatever unit it asks for.
pub const MAX_PLOT_AXIS_TICKS: usize = 64;
/// Hard ceiling on the cells one surface chart may band.
pub const MAX_PLOT_SURFACE_CELLS: usize = 16_384;
/// Hard ceiling on the per-point `c:dLbl` overrides one series may carry.
pub const MAX_PLOT_DATA_LABELS: usize = 4_096;
/// Hard ceiling on the vertices one area or radar polygon may carry.
pub const MAX_PLOT_POLYGON_POINTS: usize = 8_192;
/// Coordinates are clamped here so every emitted op stays finite.
pub const MAX_PLOT_COORD: f64 = 1e9;

const MAX_LABEL_CHARS: usize = 120;
const MAX_LEGEND_ENTRIES: usize = 8;

#[derive(Clone, Debug, PartialEq)]
pub struct PlotFont {
    pub weight: u16,
    pub size_px: f64,
    pub family: String,
    pub italic: bool,
}

impl PlotFont {
    /// CSS `font` shorthand, for hosts that paint through a browser.
    pub fn css(&self) -> String {
        let style = if self.italic { "italic " } else { "" };
        format!("{style}{} {}px {}", self.weight, self.size_px, self.family)
    }
}

/// One `c:txPr`, every field optional so an unset one inherits.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotTextStyle<'a> {
    pub font: Option<&'a str>,
    pub size_pt: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<&'a str>,
}

impl<'a> PlotTextStyle<'a> {
    /// `self` over `base`, for the chart-scope style a scoped one inherits.
    fn over(self, base: PlotTextStyle<'a>) -> Self {
        Self {
            font: self.font.or(base.font),
            size_pt: self.size_pt.or(base.size_pt),
            bold: self.bold.or(base.bold),
            italic: self.italic.or(base.italic),
            color: self.color.or(base.color),
        }
    }

    /// The style resolved against a default size and weight.
    fn resolve(self, size_px: f64, weight: u16) -> ResolvedText {
        ResolvedText {
            font: PlotFont {
                weight: match self.bold {
                    Some(true) => 700,
                    Some(false) => 400,
                    None => weight,
                },
                size_px: self
                    .size_pt
                    .filter(|size| size.is_finite() && *size > 0.0)
                    .map(|size| (size * 4.0 / 3.0).clamp(1.0, 400.0))
                    .unwrap_or(size_px),
                family: self
                    .font
                    .filter(|font| !font.is_empty())
                    .map(str::to_owned)
                    .unwrap_or_else(|| CHART_FONT_FAMILY.to_owned()),
                italic: self.italic.unwrap_or(false),
            },
            color: self
                .color
                .filter(|color| !color.is_empty())
                .map(hex)
                .unwrap_or_else(|| CHART_TEXT_COLOR.to_owned()),
        }
    }
}

/// What one text op paints with.
#[derive(Clone, Debug, PartialEq)]
struct ResolvedText {
    font: PlotFont,
    color: String,
}

/// Chart-scope `c:txPr` plus the scopes that inherit from it.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotChartText<'a> {
    pub chart: PlotTextStyle<'a>,
    pub title: PlotTextStyle<'a>,
    pub legend: PlotTextStyle<'a>,
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

/// Receives plot ops in order and returns whether plotting may continue.
pub trait PlotSink {
    fn accepts_more(&mut self) -> bool {
        true
    }

    fn push_op(&mut self, op: PlotOp) -> bool;
}

impl PlotSink for Vec<PlotOp> {
    fn push_op(&mut self, op: PlotOp) -> bool {
        self.push(op);
        true
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
    /// `c:txPr` at chart, title and legend scope.
    pub text: PlotChartText<'a>,
    /// Every `c:catAx`, `c:valAx`, `c:dateAx` and `c:serAx` of the plot area,
    /// in document order, so a plot group can find the axis it names.
    pub axes: Vec<PlotAxis<'a>>,
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

/// Which family of values an axis carries; charts key their scales off it.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum PlotAxisKind {
    #[default]
    Category,
    Value,
    Date,
    Series,
}

impl PlotAxisKind {
    pub fn from_name(name: &str) -> Self {
        match name {
            "value" => Self::Value,
            "date" => Self::Date,
            "series" => Self::Series,
            _ => Self::Category,
        }
    }
}

/// One `c:catAx`, `c:valAx`, `c:dateAx` or `c:serAx`.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotAxis<'a> {
    pub id: Option<&'a str>,
    pub kind: PlotAxisKind,
    pub range: PlotAxisRange,
    pub log_base: Option<f64>,
    pub reversed: bool,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    pub major_tick_mark: Option<&'a str>,
    pub minor_tick_mark: Option<&'a str>,
    pub major_gridlines: bool,
    pub minor_gridlines: bool,
    pub number_format: Option<&'a str>,
    pub position: Option<&'a str>,
    pub title: Option<&'a str>,
    pub hidden: bool,
    /// `c:txPr` on this axis.
    pub text: PlotTextStyle<'a>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PlotGroup<'a> {
    pub chart_type: Option<&'a str>,
    pub grouping: Option<&'a str>,
    pub series: Vec<PlotSeries<'a>>,
    pub overlap: Option<f64>,
    pub gap_width: Option<f64>,
    pub hole_size: Option<f64>,
    pub first_slice_angle: Option<f64>,
    pub vary_colors: bool,
    pub scatter_style: Option<&'a str>,
    pub radar_style: Option<&'a str>,
    pub bubble_scale: Option<f64>,
    pub size_represents: Option<&'a str>,
    pub wireframe: Option<bool>,
    pub hi_low_lines: bool,
    pub up_down_bars: bool,
    /// `c:marker` of a line chart: `Some(false)` switches every marker off.
    pub markers: Option<bool>,
    pub axis_ids: Vec<&'a str>,
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
    /// `c:xVal` of a scatter or bubble series; empty means categorical x.
    pub x_values: &'a [f64],
    /// `c:bubbleSize` of a bubble series.
    pub bubble_sizes: &'a [f64],
    pub smooth: bool,
    /// This series' `c:dLbls`, already merged over its plot group's.
    pub labels: Option<PlotDataLabels<'a>>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotPoint<'a> {
    pub index: Option<usize>,
    pub value: Option<f64>,
    pub color: Option<&'a str>,
    pub marker: Option<PlotMarker>,
    /// Literal label text, which wins over anything [`PlotDataLabels`] would
    /// compose.
    pub label: Option<&'a str>,
    /// `c:explosion`, a percentage of the pie radius.
    pub explosion: Option<f64>,
    /// This point's cascade-resolved `c:dLbl`.
    pub labels: Option<PlotDataLabels<'a>>,
}

/// A resolved `c:dLbls`: which fields a label shows and where it sits.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotDataLabels<'a> {
    pub show_value: bool,
    pub show_category_name: bool,
    pub show_series_name: bool,
    pub show_percent: bool,
    pub show_legend_key: bool,
    pub show_bubble_size: bool,
    pub separator: Option<&'a str>,
    /// `c:dLblPos`: `ctr`, `inEnd`, `inBase`, `outEnd`, `bestFit`, `l`, `r`,
    /// `t` or `b`.
    pub position: Option<&'a str>,
    pub number_format: Option<&'a str>,
    /// `c:txPr` on this `c:dLbls`.
    pub text: PlotTextStyle<'a>,
}

impl PlotDataLabels<'_> {
    /// Whether any field would compose into text.
    pub fn shows_anything(&self) -> bool {
        self.show_value
            || self.show_category_name
            || self.show_series_name
            || self.show_percent
            || self.show_bubble_size
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PlotMarker {
    pub size: Option<f64>,
    pub symbol: Option<PlotMarkerSymbol>,
}

/// A `c:symbol`. An unset symbol keeps the square the geometry has always
/// drawn; `Auto` cycles Excel's automatic sequence by series index.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PlotMarkerSymbol {
    Auto,
    None,
    Circle,
    Dash,
    Diamond,
    Dot,
    Plus,
    Square,
    Star,
    Triangle,
    X,
}

impl PlotMarkerSymbol {
    pub fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "auto" => Self::Auto,
            "none" => Self::None,
            "circle" => Self::Circle,
            "dash" => Self::Dash,
            "diamond" => Self::Diamond,
            "dot" => Self::Dot,
            "plus" => Self::Plus,
            "square" => Self::Square,
            "star" => Self::Star,
            "triangle" => Self::Triangle,
            "x" => Self::X,
            _ => return None,
        })
    }

    /// What `auto` resolves to for the `index`th series.
    fn resolved(self, index: usize) -> Self {
        const CYCLE: [PlotMarkerSymbol; 8] = [
            PlotMarkerSymbol::Diamond,
            PlotMarkerSymbol::Square,
            PlotMarkerSymbol::Triangle,
            PlotMarkerSymbol::X,
            PlotMarkerSymbol::Star,
            PlotMarkerSymbol::Dot,
            PlotMarkerSymbol::Plus,
            PlotMarkerSymbol::Dash,
        ];
        match self {
            Self::Auto => CYCLE[index % CYCLE.len()],
            symbol => symbol,
        }
    }
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
            series: space
                .series
                .iter()
                .map(|series| plot_series_from_model(series, None))
                .collect(),
            plot_groups: space
                .plot_groups
                .iter()
                .map(|group| PlotGroup {
                    chart_type: group.chart_type.as_deref(),
                    grouping: group.grouping.as_deref(),
                    series: group
                        .series
                        .iter()
                        .map(|series| plot_series_from_model(series, group.data_labels.as_ref()))
                        .collect(),
                    overlap: group.overlap,
                    gap_width: group.gap_width,
                    hole_size: group.hole_size,
                    first_slice_angle: group.first_slice_angle,
                    vary_colors: group.vary_colors,
                    scatter_style: group.scatter_style.as_deref(),
                    radar_style: group.radar_style.as_deref(),
                    bubble_scale: group.bubble_scale,
                    size_represents: group.size_represents.as_deref(),
                    wireframe: group.wireframe,
                    hi_low_lines: group.hi_low_lines,
                    up_down_bars: group.up_down_bars,
                    markers: group.marker,
                    axis_ids: group.axis_ids.iter().map(String::as_str).collect(),
                })
                .collect(),
            axes: space
                .axis_list
                .iter()
                .flatten()
                .map(plot_axis_from_model)
                .collect(),
            text: PlotChartText {
                chart: plot_text_from_model(space.text.as_ref()),
                title: plot_text_from_model(space.title_text.as_ref()),
                legend: plot_text_from_model(
                    space
                        .legend
                        .as_ref()
                        .and_then(|legend| legend.text.as_ref()),
                ),
            },
        }
    }
}

fn plot_text_from_model(text: Option<&super::model::ChartTextProperties>) -> PlotTextStyle<'_> {
    text.map(|text| PlotTextStyle {
        font: text.font.as_deref(),
        size_pt: text.size_pt,
        bold: text.bold,
        italic: text.italic,
        color: text.color.as_deref(),
    })
    .unwrap_or_default()
}

fn plot_axis_from_model(axis: &super::model::ChartAxis) -> PlotAxis<'_> {
    PlotAxis {
        id: axis.id.as_deref(),
        kind: PlotAxisKind::from_name(&axis.axis_type),
        range: PlotAxisRange {
            min: axis.min,
            max: axis.max,
        },
        log_base: axis.logarithmic_base,
        reversed: axis.reversed,
        major_unit: axis.major_unit,
        minor_unit: axis.minor_unit,
        major_tick_mark: axis.major_tick_mark.as_deref(),
        minor_tick_mark: axis.minor_tick_mark.as_deref(),
        major_gridlines: axis.major_gridlines,
        minor_gridlines: axis.minor_gridlines,
        number_format: axis.number_format.as_deref(),
        position: axis.position.as_deref(),
        title: axis.title.as_deref(),
        hidden: axis.hidden,
        text: plot_text_from_model(axis.text.as_ref()),
    }
}

fn plot_series_from_model<'a>(
    series: &'a super::model::ChartSeries,
    group_labels: Option<&'a super::model::ChartDataLabels>,
) -> PlotSeries<'a> {
    let labels = plot_labels_from_model(None, series.data_labels.as_ref(), None, group_labels);
    let mut points: Vec<PlotPoint<'a>> = series
        .points
        .iter()
        .flatten()
        .filter_map(|point| {
            Some(PlotPoint {
                index: match point.index {
                    Some(index) => Some(point_index(index)?),
                    None => None,
                },
                color: Some(&point.color),
                explosion: point.explosion,
                labels,
                ..PlotPoint::default()
            })
        })
        .collect();
    merge_point_labels(&mut points, group_labels, series.data_labels.as_ref());
    PlotSeries {
        name: series.name.as_deref(),
        categories: &series.categories,
        values: &series.values,
        color: Some(&series.color),
        points,
        grouping: series.grouping.as_deref(),
        marker: series.marker.as_ref().map(plot_marker_from_model),
        x_values: series.x_values.as_deref().unwrap_or_default(),
        bubble_sizes: series.bubble_sizes.as_deref().unwrap_or_default(),
        smooth: series.smooth.unwrap_or(false),
        labels,
    }
}

/// Resolves the four-level `c:dLbls` cascade.
fn plot_labels_from_model<'a>(
    series_point: Option<&'a super::model::ChartDataLabels>,
    series: Option<&'a super::model::ChartDataLabels>,
    group_point: Option<&'a super::model::ChartDataLabels>,
    group: Option<&'a super::model::ChartDataLabels>,
) -> Option<PlotDataLabels<'a>> {
    let levels = [series_point, series, group_point, group];
    levels.iter().copied().flatten().next()?;
    let flag = |read: fn(&super::model::ChartDataLabels) -> Option<bool>| {
        levels
            .iter()
            .copied()
            .find_map(|labels| labels.and_then(read))
    };
    if flag(|labels| labels.delete) == Some(true) {
        return None;
    }
    let switch =
        |read: fn(&super::model::ChartDataLabels) -> Option<bool>| flag(read).unwrap_or(false);
    let text = |read: fn(&super::model::ChartDataLabels) -> Option<&str>| {
        levels
            .iter()
            .copied()
            .find_map(|labels| labels.and_then(read))
    };
    let mut style = PlotTextStyle::default();
    for labels in levels.iter().rev().flatten() {
        style = plot_text_from_model(labels.text.as_ref()).over(style);
    }
    Some(PlotDataLabels {
        show_value: switch(|labels| labels.show_value),
        show_category_name: switch(|labels| labels.show_category_name),
        show_series_name: switch(|labels| labels.show_series_name),
        show_percent: switch(|labels| labels.show_percent),
        show_legend_key: switch(|labels| labels.show_legend_key),
        show_bubble_size: switch(|labels| labels.show_bubble_size),
        separator: text(|labels| labels.separator.as_deref()),
        position: text(|labels| labels.position.as_deref()),
        number_format: text(|labels| labels.number_format.as_deref()),
        text: style,
    })
}

/// Folds cascade-resolved `c:dLbl` overrides into `points`.
fn merge_point_labels<'a>(
    points: &mut Vec<PlotPoint<'a>>,
    group: Option<&'a super::model::ChartDataLabels>,
    series: Option<&'a super::model::ChartDataLabels>,
) {
    let group_overrides = group
        .and_then(|labels| labels.points.as_deref())
        .unwrap_or_default();
    let series_overrides = series
        .and_then(|labels| labels.points.as_deref())
        .unwrap_or_default();
    let mut group_points = HashMap::new();
    for over in group_overrides.iter().take(MAX_PLOT_DATA_LABELS) {
        if let Some(index) = over.index.and_then(point_index) {
            group_points.entry(index).or_insert(over);
        }
    }
    let mut series_points = HashSet::new();
    let wildcard = points.iter().position(|point| point.index.is_none());
    for over in series_overrides.iter().take(MAX_PLOT_DATA_LABELS) {
        let Some(index) = over.index.and_then(point_index) else {
            continue;
        };
        series_points.insert(index);
        merge_point_label(
            points,
            wildcard,
            index,
            group_points.get(&index).copied(),
            Some(over),
            group,
            series,
        );
    }
    let mut merged_group_points = HashSet::new();
    for over in group_overrides.iter().take(MAX_PLOT_DATA_LABELS) {
        let Some(index) = over.index.and_then(point_index) else {
            continue;
        };
        if !series_points.contains(&index) && merged_group_points.insert(index) {
            merge_point_label(points, wildcard, index, Some(over), None, group, series);
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn merge_point_label<'a>(
    points: &mut Vec<PlotPoint<'a>>,
    wildcard: Option<usize>,
    index: usize,
    group_point: Option<&'a super::model::ChartPointLabel>,
    series_point: Option<&'a super::model::ChartPointLabel>,
    group: Option<&'a super::model::ChartDataLabels>,
    series: Option<&'a super::model::ChartDataLabels>,
) {
    let resolved = plot_labels_from_model(
        series_point.map(|point| &point.labels),
        series,
        group_point.map(|point| &point.labels),
        group,
    );
    let label = resolved.and_then(|_| {
        series_point
            .and_then(|point| point.text.as_deref())
            .or_else(|| group_point.and_then(|point| point.text.as_deref()))
    });
    let labels = Some(resolved.unwrap_or_default());
    if let Some(slot) = points
        .iter()
        .position(|point| point.index == Some(index))
        .filter(|slot| wildcard.is_none_or(|wildcard| *slot <= wildcard))
    {
        points[slot].label = label;
        points[slot].labels = labels;
        return;
    }
    let mut point = PlotPoint {
        index: Some(index),
        label,
        labels,
        ..PlotPoint::default()
    };
    match wildcard {
        Some(wildcard) => {
            point.color = points[wildcard].color;
            point.marker = points[wildcard].marker;
            points.insert(wildcard, point);
        }
        None => points.push(point),
    }
}

fn plot_marker_from_model(marker: &super::model::ChartMarker) -> PlotMarker {
    PlotMarker {
        size: marker.size,
        symbol: marker
            .symbol
            .as_deref()
            .and_then(PlotMarkerSymbol::from_name),
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
    let chart_text = chart.text.chart;
    let label_style = &chart_text.resolve(CHART_LABEL_SIZE_PX, 400);
    let legend_style = &chart
        .text
        .legend
        .over(chart_text)
        .resolve(CHART_LABEL_SIZE_PX, 400);

    push_rect(ops, x, y, width, height, CHART_BACKGROUND_COLOR);

    let title_h = if let Some(title) = chart.title.filter(|s| !s.is_empty()) {
        push_text(
            ops,
            title,
            x + 8.0,
            y + 18.0,
            (width - 16.0).max(0.0),
            &chart
                .text
                .title
                .over(chart_text)
                .resolve(CHART_TITLE_SIZE_PX, 600),
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
    let secondary_w = if secondary_value_axis(chart).is_some() {
        38.0
    } else {
        0.0
    };
    let plot = PlotArea {
        x: plot_x,
        y: y + title_h,
        w: (width - 42.0 - legend_w - 10.0 - secondary_w).max(24.0),
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
                group: None,
                chart_text,
                label: label_style,
                axis: None,
                x_axis: None,
                category_axis: chart
                    .axes
                    .iter()
                    .find(|axis| axis.kind != PlotAxisKind::Value),
                secondary: false,
            },
            plot,
            x,
            y + title_h,
            width,
            height - title_h,
        );
    } else {
        let primary = primary_value_axis(chart);
        for group in chart.plot_groups.iter().take(MAX_PLOT_GROUPS) {
            if ops.exhausted() {
                break;
            }
            let series = series_views(&group.series, scan);
            let (x_axis, axis) = group_value_axes(chart, group);
            let category_axis = group_category_axis(chart, group);
            let legacy_axes = chart.axes.is_empty() && group.axis_ids.is_empty();
            let secondary = match (axis, primary) {
                (Some(axis), Some(primary)) => axis.id != primary.id,
                _ => false,
            };
            emit_family(
                ops,
                PlotFamily {
                    chart_type: group.chart_type.unwrap_or(chart.chart_type),
                    series: &series,
                    value_axis: legacy_axes.then_some(chart.value_axis).flatten(),
                    axis_titles: if legacy_axes {
                        chart.axis_titles
                    } else {
                        PlotAxisTitles {
                            category: category_axis.and_then(|axis| axis.title),
                            value: axis.and_then(|axis| axis.title),
                        }
                    },
                    group: Some(group),
                    chart_text,
                    label: label_style,
                    axis,
                    x_axis,
                    category_axis,
                    secondary,
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
        legend_style,
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
        let family = chart
            .plot_groups
            .first()
            .and_then(|group| group.chart_type)
            .unwrap_or(chart.chart_type);
        match family {
            "bar" => "bar chart",
            "line" => "line chart",
            "pie" | "ofPie" => "pie chart",
            "doughnut" => "doughnut chart",
            "area" => "area chart",
            "scatter" => "scatter chart",
            "bubble" => "bubble chart",
            "radar" => "radar chart",
            "stock" => "stock chart",
            "surface" => "surface chart",
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
        if self.exhausted() {
            return;
        }
        if self.sink.push_op(op) {
            self.remaining -= 1;
        } else {
            self.remaining = 0;
        }
    }

    fn exhausted(&mut self) -> bool {
        self.remaining == 0 || !self.sink.accepts_more()
    }
}

/// The slice of a chart one plot family draws: a combo chart emits one per
/// plot group, each against whichever value axis its `c:axId` list names.
#[derive(Clone, Copy)]
struct PlotFamily<'a> {
    chart_type: &'a str,
    series: &'a [SeriesView<'a>],
    value_axis: Option<PlotAxisRange>,
    axis_titles: PlotAxisTitles<'a>,
    group: Option<&'a PlotGroup<'a>>,
    /// Chart-scope `c:txPr`, which every scoped style inherits from.
    chart_text: PlotTextStyle<'a>,
    /// Chart-scope text already resolved, for labels with no scope of their own.
    label: &'a ResolvedText,
    axis: Option<&'a PlotAxis<'a>>,
    /// The x value axis of a scatter or bubble group.
    x_axis: Option<&'a PlotAxis<'a>>,
    category_axis: Option<&'a PlotAxis<'a>>,
    /// The group plots against a value axis other than the chart's first.
    secondary: bool,
}

impl<'a> PlotFamily<'a> {
    fn label(&self) -> &ResolvedText {
        self.label
    }

    /// `scope` resolved over the chart's own `c:txPr`.
    fn scoped(&self, scope: PlotTextStyle<'a>) -> ResolvedText {
        scope
            .over(self.chart_text)
            .resolve(CHART_LABEL_SIZE_PX, 400)
    }

    /// The style the category axis gives its own tick labels.
    fn category_text(&self) -> ResolvedText {
        self.scoped(self.category_axis.map(|axis| axis.text).unwrap_or_default())
    }

    fn grouping(&self) -> &'a str {
        self.group
            .and_then(|group| group.grouping)
            .or_else(|| {
                self.series
                    .first()
                    .and_then(|series| series.series.grouping)
            })
            .unwrap_or("clustered")
    }

    fn stacking(&self) -> Stacking {
        match self.grouping() {
            "stacked" => Stacking::Stacked,
            "percentStacked" => Stacking::Percent,
            _ => Stacking::None,
        }
    }
}

/// How a family piles its series onto one another.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Stacking {
    None,
    Stacked,
    Percent,
}

fn unique_axis_by_id<'a>(chart: &'a PlotChart<'a>, id: &str) -> Option<&'a PlotAxis<'a>> {
    let mut matches = chart.axes.iter().filter(|axis| axis.id == Some(id));
    let axis = matches.next()?;
    matches.next().is_none().then_some(axis)
}

fn referenced_axis<'a>(
    chart: &'a PlotChart<'a>,
    id: &str,
    kind: PlotAxisKind,
) -> Option<&'a PlotAxis<'a>> {
    unique_axis_by_id(chart, id).filter(|axis| axis.kind == kind)
}

/// Resolves a group's value-axis references as `(x, y)`.
fn group_value_axes<'a>(
    chart: &'a PlotChart<'a>,
    group: &PlotGroup<'a>,
) -> (Option<&'a PlotAxis<'a>>, Option<&'a PlotAxis<'a>>) {
    if matches!(
        group.chart_type.unwrap_or(chart.chart_type),
        "scatter" | "bubble"
    ) {
        let x_id = group.axis_ids.first().copied();
        let y_id = group.axis_ids.get(1).copied();
        if x_id.is_some() && x_id == y_id {
            return (None, None);
        }
        return (
            x_id.and_then(|id| referenced_axis(chart, id, PlotAxisKind::Value)),
            y_id.and_then(|id| referenced_axis(chart, id, PlotAxisKind::Value)),
        );
    }
    let mut named = group
        .axis_ids
        .iter()
        .filter_map(|id| referenced_axis(chart, id, PlotAxisKind::Value));
    match (named.next(), named.next()) {
        (Some(axis), None) => (None, Some(axis)),
        _ => (None, None),
    }
}

/// The value axis the chart's first plot group measures against.
fn primary_value_axis<'a>(chart: &'a PlotChart<'a>) -> Option<&'a PlotAxis<'a>> {
    chart
        .plot_groups
        .first()
        .and_then(|group| group_value_axes(chart, group).1)
}

fn group_category_axis<'a>(
    chart: &'a PlotChart<'a>,
    group: &PlotGroup<'a>,
) -> Option<&'a PlotAxis<'a>> {
    let mut named = group
        .axis_ids
        .iter()
        .filter_map(|id| unique_axis_by_id(chart, id))
        .filter(|axis| matches!(axis.kind, PlotAxisKind::Category | PlotAxisKind::Date));
    match (named.next(), named.next()) {
        (Some(axis), None) => Some(axis),
        _ => None,
    }
}

/// A second value axis some plot group plots against, which needs its own
/// scale and its own labels on the far side of the plot area.
fn secondary_value_axis<'a>(chart: &'a PlotChart<'a>) -> Option<&'a PlotAxis<'a>> {
    let primary = primary_value_axis(chart)?;
    chart.plot_groups.iter().find_map(|group| {
        let axis = group_value_axes(chart, group).1?;
        (axis.id != primary.id).then_some(axis)
    })
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
        self.data_value(index).unwrap_or(0.0)
    }

    fn data_value(&self, index: usize) -> Option<f64> {
        self.point(index)
            .and_then(|point| point.value)
            .or_else(|| self.series.values.get(index).copied())
            .filter(|value| value.is_finite())
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

    /// The symbol to draw at `index`, or `None` for a point that draws none.
    fn marker_symbol(&self, index: usize, series_index: usize) -> Option<PlotMarkerSymbol> {
        let symbol = self
            .point(index)
            .and_then(|point| point.marker.as_ref())
            .or(self.series.marker.as_ref())
            .and_then(|marker| marker.symbol)
            .unwrap_or(PlotMarkerSymbol::Square)
            .resolved(series_index);
        (symbol != PlotMarkerSymbol::None).then_some(symbol)
    }

    fn x_value(&self, index: usize) -> Option<f64> {
        self.series
            .x_values
            .get(index)
            .copied()
            .filter(|value| value.is_finite())
    }

    fn xy_value(&self, index: usize) -> Option<(f64, f64)> {
        Some((self.x_value(index)?, self.data_value(index)?))
    }

    fn bubble_size(&self, index: usize) -> f64 {
        self.series
            .bubble_sizes
            .get(index)
            .copied()
            .filter(|value| value.is_finite() && *value > 0.0)
            .unwrap_or(0.0)
    }

    fn explosion(&self, index: usize) -> f64 {
        self.point(index)
            .and_then(|point| point.explosion)
            .filter(|value| value.is_finite())
            .unwrap_or(0.0)
            .clamp(0.0, 400.0)
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
        "line" => emit_line(ops, family, plot),
        "area" => emit_area(ops, family, plot),
        "scatter" => emit_scatter(ops, family, plot),
        "bubble" => emit_bubble(ops, family, plot),
        "radar" => emit_radar(ops, family, plot, x, y, width, height),
        "stock" => emit_stock(ops, family, plot),
        "surface" => emit_surface(ops, family, plot),
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
    if w <= 0.0 || h <= 0.0 || ops.exhausted() {
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
    style: &ResolvedText,
) {
    if text.is_empty() || width <= 0.0 || ops.exhausted() {
        return;
    }
    ops.push(PlotOp::Text {
        text: text.chars().take(MAX_LABEL_CHARS).collect(),
        x,
        baseline_y,
        width,
        font: style.font.clone(),
        color: style.color.clone(),
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
    if ops.exhausted() {
        return;
    }
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
        .max(series.x_values.len())
        .max(series.bubble_sizes.len())
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

/// Maps a value onto the plot area, honouring log bases and reversed axes.
#[derive(Clone, Copy, Debug, PartialEq)]
struct ValueScale {
    min: f64,
    max: f64,
    log_base: Option<f64>,
    reversed: bool,
    percent: bool,
}

impl ValueScale {
    /// Where `value` sits between the bounds, 0 at the axis start and 1 at its
    /// end, before any reversal.
    fn fraction(&self, value: f64) -> f64 {
        let raw = match self.log_base {
            Some(base) if base > 1.0 && self.min > 0.0 && self.max > self.min => {
                let log = |value: f64| value.max(f64::MIN_POSITIVE).log(base);
                (log(value.max(self.min)) - log(self.min)) / (log(self.max) - log(self.min))
            }
            _ => (value - self.min) / (self.max - self.min),
        };
        if raw.is_finite() { raw } else { 0.0 }
    }

    fn ratio(&self, value: f64) -> f64 {
        let fraction = self.fraction(value).clamp(0.0, 1.0);
        if self.reversed {
            1.0 - fraction
        } else {
            fraction
        }
    }

    fn y(&self, plot: PlotArea, value: f64) -> f64 {
        plot.y + (1.0 - self.ratio(value)) * plot.h
    }

    fn x(&self, plot: PlotArea, value: f64) -> f64 {
        plot.x + self.ratio(value) * plot.w
    }

    /// The value the axis carries zero at, pulled inside the bounds.
    fn baseline(&self) -> f64 {
        0.0_f64.clamp(self.min, self.max)
    }

    fn format(&self, value: f64, code: Option<&str>) -> String {
        if let Some(formatted) = code.and_then(|code| format_with_code(value, code)) {
            return formatted;
        }
        if self.percent {
            return format_percent(value);
        }
        format_number(value)
    }
}

/// Positive and negative running totals per category, for a stacked family.
fn stacked_totals(family: PlotFamily<'_>) -> (f64, f64) {
    let (mut low, mut high) = (0.0_f64, 0.0_f64);
    let categories = category_count(family.series).min(MAX_PLOT_DATA_SCAN);
    for index in 0..categories {
        let (mut negative, mut positive) = (0.0_f64, 0.0_f64);
        for series in family.series {
            let value = series.value(index);
            if !value.is_finite() {
                continue;
            }
            if value < 0.0 {
                negative += value;
            } else {
                positive += value;
            }
        }
        low = low.min(negative);
        high = high.max(positive);
    }
    (low, high)
}

/// A percent-stacked family reaches 100% upward and, when a category carries
/// negative values, their share downward.
fn percent_range(family: PlotFamily<'_>) -> (f64, f64) {
    let (mut low, mut high) = (0.0_f64, 0.0_f64);
    let categories = category_count(family.series).min(MAX_PLOT_DATA_SCAN);
    for index in 0..categories {
        let (mut negative, mut total) = (0.0_f64, 0.0_f64);
        for series in family.series {
            let value = series.value(index);
            if !value.is_finite() {
                continue;
            }
            total += value.abs();
            if value < 0.0 {
                negative += value;
            }
        }
        if total > 0.0 {
            low = low.min(negative / total);
            high = high.max(1.0 + negative / total);
        }
    }
    (low, if high > 0.0 { high } else { 1.0 })
}

fn plain_range(family: PlotFamily<'_>) -> (f64, f64) {
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
    (min, max)
}

#[cfg(test)]
fn value_range(family: PlotFamily<'_>) -> (f64, f64) {
    let scale = value_scale(family);
    (scale.min, scale.max)
}

fn value_scale(family: PlotFamily<'_>) -> ValueScale {
    let stacking = family.stacking();
    let (mut min, mut max) = match stacking {
        Stacking::Percent => percent_range(family),
        Stacking::Stacked => stacked_totals(family),
        Stacking::None => plain_range(family),
    };
    let bounds = family
        .axis
        .map(|axis| axis.range)
        .or(family.value_axis)
        .unwrap_or_default();
    if let Some(value) = bounds.min.filter(|value| value.is_finite()) {
        min = value;
    }
    if let Some(value) = bounds.max.filter(|value| value.is_finite()) {
        max = value;
    }
    if max <= min {
        max = min + 1.0;
    }
    if !(max - min).is_finite() || max <= min {
        (min, max) = (0.0, 1.0);
    }
    ValueScale {
        min,
        max,
        log_base: family
            .axis
            .and_then(|axis| axis.log_base)
            .filter(|base| *base > 1.0 && base.is_finite() && min > 0.0),
        reversed: family.axis.is_some_and(|axis| axis.reversed),
        percent: stacking == Stacking::Percent,
    }
}

/// The major tick values of `scale`: powers of the base on a log axis, at
/// `c:majorUnit` when it names one, and at five even steps otherwise.
fn axis_ticks(scale: ValueScale, unit: Option<f64>) -> Vec<f64> {
    let span = scale.max - scale.min;
    if let Some(base) = scale.log_base {
        let first = scale.min.log(base).ceil();
        let last = scale.max.log(base).floor();
        if last >= first && (last - first) < MAX_PLOT_AXIS_TICKS as f64 {
            let ticks: Vec<f64> = (0..=(last - first) as usize)
                .map(|step| base.powf(first + step as f64))
                .filter(|value| value.is_finite())
                .collect();
            if ticks.len() >= 2 {
                return ticks;
            }
        }
    }
    if let Some(unit) = unit.filter(|unit| unit.is_finite() && *unit > 0.0) {
        let steps = (span / unit).floor();
        if steps >= 1.0 && steps < MAX_PLOT_AXIS_TICKS as f64 {
            let first = (scale.min / unit).ceil() * unit;
            let mut ticks = Vec::new();
            let mut index = 0;
            while ticks.len() < MAX_PLOT_AXIS_TICKS {
                let value = first + unit * index as f64;
                if !value.is_finite() || value > scale.max + unit * 1e-9 {
                    break;
                }
                ticks.push(value);
                index += 1;
            }
            if ticks.len() >= 2 {
                return ticks;
            }
        }
    }
    (0..=4)
        .map(|step| match step {
            4 => scale.max,
            step => scale.min + span * step as f64 / 4.0,
        })
        .collect()
}

/// Half-length of a tick mark drawn for `mark`, and whether it crosses.
fn tick_extents(mark: Option<&str>) -> Option<(f64, f64)> {
    match mark? {
        "in" => Some((0.0, 4.0)),
        "out" => Some((4.0, 0.0)),
        "cross" => Some((3.0, 3.0)),
        _ => None,
    }
}

fn emit_axes<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let scale = value_scale(family);
    let axis = family.axis;
    let hidden = axis.is_some_and(|axis| axis.hidden);
    let major_grid = axis.is_none_or(|axis| axis.major_gridlines);
    let minor_grid = axis.is_some_and(|axis| axis.minor_gridlines);
    let number_format = axis.and_then(|axis| axis.number_format);
    let tick_style = &family.scoped(axis.map(|axis| axis.text).unwrap_or_default());
    let (edge, outward) = if family.secondary {
        (plot.x + plot.w, 1.0)
    } else {
        (plot.x, -1.0)
    };
    let label_x = if family.secondary {
        plot.x + plot.w + 4.0
    } else {
        plot.x - 38.0
    };

    if let Some(minor_unit) = axis.and_then(|axis| axis.minor_unit).filter(|_| minor_grid) {
        for value in axis_ticks(scale, Some(minor_unit)) {
            let y = scale.y(plot, value);
            push_line(ops, plot.x, y, plot.x + plot.w, y, CHART_GRID_COLOR, 0.25);
        }
    }
    for value in axis_ticks(scale, axis.and_then(|axis| axis.major_unit))
        .into_iter()
        .rev()
    {
        if ops.exhausted() {
            return;
        }
        let y = scale.y(plot, value);
        if major_grid {
            push_line(ops, plot.x, y, plot.x + plot.w, y, CHART_GRID_COLOR, 0.5);
        }
        if hidden {
            continue;
        }
        push_text(
            ops,
            &scale.format(value, number_format),
            label_x,
            y + 3.0,
            34.0,
            tick_style,
        );
        if let Some((outer, inner)) = tick_extents(axis.and_then(|axis| axis.major_tick_mark)) {
            push_line(
                ops,
                edge + outward * outer,
                y,
                edge - outward * inner,
                y,
                CHART_AXIS_COLOR,
                1.0,
            );
        }
    }
    if family.secondary {
        push_line(
            ops,
            plot.x + plot.w,
            plot.y,
            plot.x + plot.w,
            plot.y + plot.h,
            CHART_AXIS_COLOR,
            1.0,
        );
        return;
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
            tick_style,
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
            tick_style,
        );
    }
}
/// What a point's data label reads, composing the `c:dLbls` fields the series
/// switched on. A literal `c:dLbl` text, or a host-injected label, wins whole.
fn point_label(
    family: PlotFamily<'_>,
    series: &SeriesView<'_>,
    index: usize,
    percent_total: f64,
) -> Option<String> {
    let point = series.point(index);
    if let Some(text) = point.and_then(|point| point.label) {
        return Some(text.to_owned());
    }
    let spec = point_label_spec(series, index)?;
    if !spec.shows_anything() {
        return None;
    }
    let separator = spec.separator.unwrap_or(", ");
    let mut parts: Vec<String> = Vec::with_capacity(4);
    if spec.show_series_name
        && let Some(name) = series.series.name
    {
        parts.push(name.to_owned());
    }
    if spec.show_category_name {
        parts.push(category_label(family.series, index));
    }
    let value = series.value(index);
    if spec.show_value {
        parts.push(
            spec.number_format
                .and_then(|code| format_with_code(value, code))
                .unwrap_or_else(|| format_number(value)),
        );
    }
    if spec.show_percent {
        parts.push(if percent_total > 0.0 {
            format_percent(value / percent_total)
        } else {
            format_percent(0.0)
        });
    }
    if spec.show_bubble_size {
        parts.push(format_number(series.bubble_size(index)));
    }
    (!parts.is_empty()).then(|| parts.join(separator))
}

fn point_label_spec<'a>(series: &SeriesView<'a>, index: usize) -> Option<PlotDataLabels<'a>> {
    series
        .point(index)
        .and_then(|point| point.labels)
        .or(series.series.labels)
}

/// What `c:showPercent` divides by outside a pie: the category total across
/// every series the family draws.
fn category_total(family: PlotFamily<'_>, index: usize) -> f64 {
    family
        .series
        .iter()
        .map(|series| series.value(index).abs())
        .filter(|value| value.is_finite())
        .sum()
}

/// The label swatch `c:showLegendKey` asks for, drawn before the text.
fn push_legend_key<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    series: &SeriesView<'_>,
    series_index: usize,
    index: usize,
    x: f64,
    baseline_y: f64,
) {
    let shows = point_label_spec(series, index).is_some_and(|labels| labels.show_legend_key);
    if shows {
        push_rect(
            ops,
            x - 10.0,
            baseline_y - 7.0,
            7.0,
            7.0,
            &series.point_color(index, series_index),
        );
    }
}

/// Draws one point's data label at `(x, baseline_y)`, with its legend key.
#[allow(clippy::too_many_arguments)]
fn push_point_label<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    series: &SeriesView<'_>,
    series_index: usize,
    index: usize,
    x: f64,
    baseline_y: f64,
    width: f64,
    percent_total: f64,
) {
    let Some(text) = point_label(family, series, index, percent_total) else {
        return;
    };
    let scope = point_label_spec(series, index)
        .map(|labels| labels.text)
        .unwrap_or_default();
    push_legend_key(ops, series, series_index, index, x, baseline_y);
    push_text(ops, &text, x, baseline_y, width, &family.scoped(scope));
}

/// Where `c:dLblPos` puts a bar label, as a fraction of the bar's own span
/// measured from its base, plus the pixels it stands off the end.
fn bar_label_anchor(position: Option<&str>) -> (f64, f64) {
    match position {
        Some("ctr") => (0.5, 0.0),
        Some("inEnd") => (1.0, -3.0),
        Some("inBase") => (0.0, 3.0),
        _ => (1.0, 3.0),
    }
}

/// Where `c:dLblPos` puts a wedge label, as a fraction of the pie radius.
fn pie_label_reach(position: Option<&str>) -> f64 {
    match position {
        Some("ctr") => 0.5,
        Some("inEnd") => 0.8,
        Some("outEnd") => 1.15,
        _ => 0.62,
    }
}

/// Where one category's bars sit inside its slot, from `c:gapWidth` and
/// `c:overlap`.
#[derive(Clone, Copy)]
struct BarBands {
    slot: f64,
    bar: f64,
    lead: f64,
    step: f64,
}

fn bar_bands(family: PlotFamily<'_>, categories: usize, extent: f64) -> BarBands {
    let stacking = family.stacking();
    let lanes = if stacking == Stacking::None {
        family.series.len().max(1)
    } else {
        1
    };
    let gap = family
        .group
        .and_then(|group| group.gap_width)
        .filter(|gap| gap.is_finite())
        .unwrap_or(150.0)
        .clamp(0.0, 500.0)
        / 100.0;
    let overlap = family
        .group
        .and_then(|group| group.overlap)
        .filter(|overlap| overlap.is_finite())
        .unwrap_or(if stacking == Stacking::None {
            0.0
        } else {
            100.0
        })
        .clamp(-100.0, 100.0)
        / 100.0;
    let slot = extent / categories.max(1) as f64;
    let span = lanes as f64 - (lanes as f64 - 1.0) * overlap;
    let bar = (slot / (span + gap).max(0.1)).clamp(0.5, extent.max(0.5));
    BarBands {
        slot,
        bar,
        lead: (slot - bar * span).max(0.0) / 2.0,
        step: bar * (1.0 - overlap),
    }
}

/// Drawing position of the `index`th category, which a reversed category axis
/// counts from the far end.
fn category_position(family: PlotFamily<'_>, index: usize, count: usize) -> usize {
    if family.category_axis.is_some_and(|axis| axis.reversed) {
        count.saturating_sub(1).saturating_sub(index)
    } else {
        index
    }
}

/// The `[start, end]` value span each series occupies in one category, piled
/// up for a stacked family and measured from the baseline otherwise.
fn stacked_spans(family: PlotFamily<'_>, category: usize, spans: &mut Vec<(f64, f64)>) {
    spans.clear();
    let stacking = family.stacking();
    let total: f64 = if stacking == Stacking::Percent {
        family
            .series
            .iter()
            .map(|series| series.value(category).abs())
            .filter(|value| value.is_finite())
            .sum()
    } else {
        0.0
    };
    let (mut positive, mut negative) = (0.0_f64, 0.0_f64);
    for series in family.series {
        let mut value = series.value(category);
        if !value.is_finite() {
            value = 0.0;
        }
        if stacking == Stacking::Percent {
            value = if total > 0.0 { value / total } else { 0.0 };
        }
        match stacking {
            Stacking::None => spans.push((0.0, value)),
            _ if value < 0.0 => {
                spans.push((negative, negative + value));
                negative += value;
            }
            _ => {
                spans.push((positive, positive + value));
                positive += value;
            }
        }
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
    let scale = value_scale(family);
    let bands = bar_bands(family, cat_count, if horizontal { plot.h } else { plot.w });
    let category_style = &family.category_text();
    let spans = &mut Vec::with_capacity(family.series.len());
    for cat_idx in 0..cat_count {
        if ops.exhausted() {
            return;
        }
        let slot = bands.slot * category_position(family, cat_idx, cat_count) as f64;
        let label = category_label(family.series, cat_idx);
        if horizontal {
            push_text(
                ops,
                &label,
                plot.x - 38.0,
                plot.y + slot + bands.slot * 0.55,
                36.0,
                category_style,
            );
        } else {
            push_text(
                ops,
                &label,
                plot.x + slot + 2.0,
                plot.y + plot.h + 14.0,
                bands.slot - 4.0,
                category_style,
            );
        }
        stacked_spans(family, cat_idx, spans);
        let total = category_total(family, cat_idx);
        for (ser_idx, series) in family.series.iter().enumerate() {
            let (start, end) = spans.get(ser_idx).copied().unwrap_or((0.0, 0.0));
            let lane = if family.stacking() == Stacking::None {
                ser_idx as f64
            } else {
                0.0
            };
            let offset = slot + bands.lead + bands.step * lane;
            let color = series.point_color(cat_idx, ser_idx);
            if horizontal {
                let (x0, x1) = (scale.x(plot, start), scale.x(plot, end));
                let y = plot.y + offset;
                push_rect(ops, x0.min(x1), y, (x1 - x0).abs(), bands.bar, &color);
                let (fraction, offset) = bar_label_anchor(
                    point_label_spec(series, cat_idx).and_then(|labels| labels.position),
                );
                push_point_label(
                    ops,
                    family,
                    series,
                    ser_idx,
                    cat_idx,
                    x0 + (x1 - x0) * fraction + offset,
                    y + bands.bar,
                    48.0,
                    total,
                );
            } else {
                let (y0, y1) = (scale.y(plot, start), scale.y(plot, end));
                let x = plot.x + offset;
                push_rect(
                    ops,
                    x,
                    y0.min(y1),
                    bands.bar,
                    (y0 - y1).abs().max(1.0),
                    &color,
                );
                let (fraction, offset) = bar_label_anchor(
                    point_label_spec(series, cat_idx).and_then(|labels| labels.position),
                );
                push_point_label(
                    ops,
                    family,
                    series,
                    ser_idx,
                    cat_idx,
                    x,
                    y0 + (y1 - y0) * fraction - offset,
                    bands.bar.max(32.0),
                    total,
                );
            }
        }
    }
}

/// Where the `index`th of `count` categories sits along a line or area axis.
fn line_x(family: PlotFamily<'_>, plot: PlotArea, index: usize, count: usize) -> f64 {
    let denom = count.saturating_sub(1).max(1) as f64;
    plot.x + plot.w * category_position(family, index, count) as f64 / denom
}

fn emit_category_labels<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
    count: usize,
) {
    let style = &family.category_text();
    for index in 0..count {
        if ops.exhausted() {
            return;
        }
        push_text(
            ops,
            &category_label(family.series, index),
            line_x(family, plot, index, count) - 16.0,
            plot.y + plot.h + 14.0,
            32.0,
            style,
        );
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
    let scale = value_scale(family);
    let stacking = family.stacking();
    emit_category_labels(ops, family, plot, cat_count);
    let spans = &mut Vec::with_capacity(family.series.len());
    for (ser_idx, series) in family.series.iter().enumerate() {
        let color = series_color(Some(series.series), ser_idx);
        let markers = family.group.and_then(|group| group.markers) != Some(false);
        let mut prev: Option<(f64, f64)> = None;
        for i in 0..cat_count {
            if ops.exhausted() {
                return;
            }
            let value = if stacking == Stacking::None {
                series.value(i)
            } else {
                stacked_spans(family, i, spans);
                spans.get(ser_idx).map_or(0.0, |(_, end)| *end)
            };
            let x = line_x(family, plot, i, cat_count);
            let y = scale.y(plot, value);
            if let Some((prev_x, prev_y)) = prev {
                push_line(ops, prev_x, prev_y, x, y, &color, 2.0);
            }
            if markers {
                push_marker(
                    ops,
                    series.marker_symbol(i, ser_idx),
                    x,
                    y,
                    series.marker_size(i),
                    &series.point_color(i, ser_idx),
                );
            }
            let size = series.marker_size(i);
            push_point_label(
                ops,
                family,
                series,
                ser_idx,
                i,
                x + size,
                y - size,
                48.0,
                category_total(family, i),
            );
            prev = Some((x, y));
        }
    }
}

fn emit_area<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let cat_count = category_count(family.series);
    if cat_count == 0 || family.series.is_empty() {
        return;
    }
    emit_axes(ops, family, plot);
    let scale = value_scale(family);
    let stacking = family.stacking();
    emit_category_labels(ops, family, plot, cat_count);
    let vertices = cat_count.min(MAX_PLOT_POLYGON_POINTS);
    let spans = &mut Vec::with_capacity(family.series.len());
    for (ser_idx, series) in family.series.iter().enumerate() {
        if ops.exhausted() {
            return;
        }
        let mut upper = Vec::with_capacity(vertices);
        let mut lower = Vec::with_capacity(vertices);
        for i in 0..vertices {
            let x = line_x(family, plot, i, cat_count);
            let (base, top) = if stacking == Stacking::None {
                (scale.baseline(), series.value(i))
            } else {
                stacked_spans(family, i, spans);
                spans.get(ser_idx).copied().unwrap_or((0.0, 0.0))
            };
            upper.push((x, scale.y(plot, top)));
            lower.push((x, scale.y(plot, base)));
        }
        let mut commands = Vec::with_capacity(vertices * 2 + 2);
        for (index, (x, y)) in upper.iter().enumerate() {
            commands.push(if index == 0 {
                GeometryPathCommand::Move { x: *x, y: *y }
            } else {
                GeometryPathCommand::Line { x: *x, y: *y }
            });
        }
        for (x, y) in lower.iter().rev() {
            commands.push(GeometryPathCommand::Line { x: *x, y: *y });
        }
        commands.push(GeometryPathCommand::Close);
        let color = series_color(Some(series.series), ser_idx);
        push_path(ops, plot, commands, &color, None);
        for window in upper.windows(2) {
            if ops.exhausted() {
                return;
            }
            push_line(
                ops,
                window[0].0,
                window[0].1,
                window[1].0,
                window[1].1,
                &color,
                1.5,
            );
        }
        for (i, (x, y)) in upper.iter().enumerate() {
            push_point_label(
                ops,
                family,
                series,
                ser_idx,
                i,
                *x,
                *y - 3.0,
                48.0,
                category_total(family, i),
            );
        }
    }
}

/// The x scale of a scatter or bubble family.
fn scatter_x_scale(family: PlotFamily<'_>) -> ValueScale {
    let (mut min, mut max) = (f64::INFINITY, f64::NEG_INFINITY);
    let mut seen = false;
    let mut remaining = MAX_PLOT_DATA_SCAN;
    for series in family.series {
        let samples = series.series.x_values.len().min(remaining);
        remaining -= samples;
        for index in 0..samples {
            if let Some((value, _)) = series.xy_value(index) {
                seen = true;
                min = min.min(value);
                max = max.max(value);
            }
        }
    }
    if !seen {
        min = 0.0;
        max = 1.0;
    }
    let axis = family.x_axis;
    if let Some(value) = axis
        .and_then(|axis| axis.range.min)
        .filter(|value| value.is_finite())
    {
        min = value;
    }
    if let Some(value) = axis
        .and_then(|axis| axis.range.max)
        .filter(|value| value.is_finite())
    {
        max = value;
    }
    if !(max - min).is_finite() || max <= min {
        (min, max) = (min.min(0.0), min.min(0.0) + 1.0);
    }
    ValueScale {
        min,
        max,
        log_base: axis
            .and_then(|axis| axis.log_base)
            .filter(|base| *base > 1.0 && base.is_finite() && min > 0.0),
        reversed: axis.is_some_and(|axis| axis.reversed),
        percent: false,
    }
}

/// Ticks along the x value axis.
fn emit_scatter_x_labels<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
    scale: ValueScale,
) {
    if !family.series.iter().any(|series| {
        (0..series.series.x_values.len().min(MAX_PLOT_DATA_SCAN))
            .any(|index| series.xy_value(index).is_some())
    }) {
        return;
    }
    let format = family.x_axis.and_then(|axis| axis.number_format);
    for value in axis_ticks(scale, family.x_axis.and_then(|axis| axis.major_unit)) {
        if ops.exhausted() {
            return;
        }
        push_text(
            ops,
            &scale.format(value, format),
            scale.x(plot, value) - 16.0,
            plot.y + plot.h + 14.0,
            32.0,
            family.label(),
        );
    }
}

/// What a `c:scatterStyle` draws.
fn scatter_parts(style: Option<&str>) -> (bool, bool) {
    match style {
        Some("line" | "smooth") => (true, false),
        Some("marker") => (false, true),
        Some("none") => (false, false),
        _ => (true, true),
    }
}

fn emit_scatter<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let count = category_count(family.series);
    if count == 0 || family.series.is_empty() {
        return;
    }
    emit_axes(ops, family, plot);
    let y_scale = value_scale(family);
    let x_scale = scatter_x_scale(family);
    emit_scatter_x_labels(ops, family, plot, x_scale);
    let (lines, markers) = scatter_parts(family.group.and_then(|group| group.scatter_style));
    for (ser_idx, series) in family.series.iter().enumerate() {
        let color = series_color(Some(series.series), ser_idx);
        let mut prev: Option<(f64, f64)> = None;
        for i in 0..series.length().min(MAX_PLOT_DATA_SCAN) {
            if ops.exhausted() {
                return;
            }
            let Some((x_value, y_value)) = series.xy_value(i) else {
                prev = None;
                continue;
            };
            let x = x_scale.x(plot, x_value);
            let y = y_scale.y(plot, y_value);
            if lines && let Some((prev_x, prev_y)) = prev {
                push_line(ops, prev_x, prev_y, x, y, &color, 2.0);
            }
            if markers {
                push_marker(
                    ops,
                    series.marker_symbol(i, ser_idx),
                    x,
                    y,
                    series.marker_size(i),
                    &series.point_color(i, ser_idx),
                );
            }
            push_point_label(
                ops,
                family,
                series,
                ser_idx,
                i,
                x + 4.0,
                y - 4.0,
                48.0,
                category_total(family, i),
            );
            prev = Some((x, y));
        }
    }
}

fn emit_bubble<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let count = category_count(family.series);
    if count == 0 || family.series.is_empty() {
        return;
    }
    emit_axes(ops, family, plot);
    let y_scale = value_scale(family);
    let x_scale = scatter_x_scale(family);
    emit_scatter_x_labels(ops, family, plot, x_scale);
    let group = family.group;
    let scale_percent = group
        .and_then(|group| group.bubble_scale)
        .filter(|scale| scale.is_finite())
        .unwrap_or(100.0)
        .clamp(0.0, 300.0)
        / 100.0;
    let by_area = group.and_then(|group| group.size_represents) != Some("w");
    let largest = family
        .series
        .iter()
        .flat_map(|series| {
            (0..series.length().min(MAX_PLOT_DATA_SCAN))
                .filter_map(move |index| series.xy_value(index).map(|_| series.bubble_size(index)))
        })
        .fold(0.0_f64, f64::max);
    let max_radius = plot.w.min(plot.h) * 0.125 * scale_percent;
    for (ser_idx, series) in family.series.iter().enumerate() {
        for i in 0..series.length().min(MAX_PLOT_DATA_SCAN) {
            if ops.exhausted() {
                return;
            }
            let Some((x_value, y_value)) = series.xy_value(i) else {
                continue;
            };
            let size = series.bubble_size(i);
            if size <= 0.0 || largest <= 0.0 {
                continue;
            }
            let ratio = (size / largest).clamp(0.0, 1.0);
            let radius = (max_radius * if by_area { ratio.sqrt() } else { ratio }).max(1.0);
            let x = x_scale.x(plot, x_value);
            let y = y_scale.y(plot, y_value);
            ops.push(PlotOp::Path {
                x: x - radius,
                y: y - radius,
                w: radius * 2.0,
                h: radius * 2.0,
                commands: circle_path(x, y, radius),
                fill: series.point_color(i, ser_idx),
                stroke: Some(PlotStroke {
                    color: CHART_BACKGROUND_COLOR.to_owned(),
                    width: 1.0,
                }),
            });
            push_point_label(
                ops,
                family,
                series,
                ser_idx,
                i,
                x,
                y,
                48.0,
                category_total(family, i),
            );
        }
    }
}

fn emit_radar<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) {
    let cat_count = category_count(family.series);
    if cat_count == 0 || family.series.is_empty() {
        return;
    }
    let scale = value_scale(family);
    let spokes = cat_count.min(MAX_PLOT_POLYGON_POINTS);
    let radius = (width.min(height) * 0.34).max(6.0);
    let (cx, cy) = (x + width * 0.38, y + height * 0.5);
    let angle = |index: usize| {
        -std::f64::consts::FRAC_PI_2
            + std::f64::consts::TAU * category_position(family, index, spokes) as f64
                / spokes as f64
    };
    let at = |index: usize, value: f64| {
        let reach = radius * scale.ratio(value);
        let angle = angle(index);
        (cx + reach * angle.cos(), cy + reach * angle.sin())
    };

    for value in axis_ticks(scale, family.axis.and_then(|axis| axis.major_unit)) {
        if ops.exhausted() || scale.ratio(value) <= 0.0 {
            continue;
        }
        let ring: Vec<(f64, f64)> = (0..spokes).map(|index| at(index, value)).collect();
        for (from, to) in ring_edges(&ring) {
            push_line(ops, from.0, from.1, to.0, to.1, CHART_GRID_COLOR, 0.5);
        }
    }
    for index in 0..spokes {
        if ops.exhausted() {
            return;
        }
        let outer = at(index, scale.max);
        push_line(ops, cx, cy, outer.0, outer.1, CHART_AXIS_COLOR, 0.5);
        push_text(
            ops,
            &category_label(family.series, index),
            cx + (outer.0 - cx) * 1.1 - 16.0,
            cy + (outer.1 - cy) * 1.1,
            32.0,
            family.label(),
        );
    }

    let style = family.group.and_then(|group| group.radar_style);
    let filled = style == Some("filled");
    for (ser_idx, series) in family.series.iter().enumerate() {
        if ops.exhausted() {
            return;
        }
        let color = series_color(Some(series.series), ser_idx);
        let ring: Vec<(f64, f64)> = (0..spokes)
            .map(|index| at(index, series.value(index)))
            .collect();
        if filled {
            let mut commands = Vec::with_capacity(spokes + 1);
            for (index, (x, y)) in ring.iter().enumerate() {
                commands.push(if index == 0 {
                    GeometryPathCommand::Move { x: *x, y: *y }
                } else {
                    GeometryPathCommand::Line { x: *x, y: *y }
                });
            }
            commands.push(GeometryPathCommand::Close);
            push_path(ops, plot, commands, &color, None);
        }
        if !filled {
            for (from, to) in ring_edges(&ring) {
                if ops.exhausted() {
                    return;
                }
                push_line(ops, from.0, from.1, to.0, to.1, &color, 2.0);
            }
        }
        for (index, (x, y)) in ring.iter().enumerate() {
            if ops.exhausted() {
                return;
            }
            if style != Some("standard") && !filled {
                push_marker(
                    ops,
                    series.marker_symbol(index, ser_idx),
                    *x,
                    *y,
                    series.marker_size(index),
                    &series.point_color(index, ser_idx),
                );
            }
            push_point_label(
                ops,
                family,
                series,
                ser_idx,
                index,
                *x,
                *y - 4.0,
                48.0,
                category_total(family, index),
            );
        }
    }
}

/// The edges of a closed radar ring, without the doubled edge two spokes
/// would otherwise produce.
fn ring_edges(ring: &[(f64, f64)]) -> impl Iterator<Item = ((f64, f64), (f64, f64))> + '_ {
    let closes = ring.len() > 2;
    ring.iter()
        .enumerate()
        .filter(move |(index, _)| closes || *index + 1 < ring.len())
        .map(move |(index, point)| (*point, ring[(index + 1) % ring.len()]))
}

/// Which of a stock family's series carry open, high, low and close.
fn stock_roles(count: usize) -> Option<(Option<usize>, usize, usize, usize)> {
    match count {
        3 => Some((None, 0, 1, 2)),
        4 => Some((Some(0), 1, 2, 3)),
        n if n > 4 => Some((Some(n - 4), n - 3, n - 2, n - 1)),
        _ => None,
    }
}

fn emit_stock<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let cat_count = category_count(family.series);
    let Some((open, high, low, close)) = stock_roles(family.series.len()) else {
        emit_line(ops, family, plot);
        return;
    };
    if cat_count == 0 {
        return;
    }
    emit_axes(ops, family, plot);
    let scale = value_scale(family);
    let bands = bar_bands(family, cat_count, plot.w);
    let hi_lo = family.group.is_none_or(|group| group.hi_low_lines) || open.is_none();
    let up_down = open.is_some() && family.group.is_none_or(|group| group.up_down_bars);
    for cat_idx in 0..cat_count {
        if ops.exhausted() {
            return;
        }
        let slot = bands.slot * category_position(family, cat_idx, cat_count) as f64;
        let center = plot.x + slot + bands.slot / 2.0;
        push_text(
            ops,
            &category_label(family.series, cat_idx),
            plot.x + slot + 2.0,
            plot.y + plot.h + 14.0,
            bands.slot - 4.0,
            family.label(),
        );
        let value = |index: usize| family.series[index].value(cat_idx);
        let (high_y, low_y) = (scale.y(plot, value(high)), scale.y(plot, value(low)));
        if hi_lo {
            push_line(ops, center, high_y, center, low_y, CHART_AXIS_COLOR, 1.0);
        }
        let tick = (bands.slot * 0.25).clamp(1.0, 12.0);
        match open {
            Some(open) if up_down => {
                let (open_y, close_y) = (scale.y(plot, value(open)), scale.y(plot, value(close)));
                let rising = value(close) >= value(open);
                push_rect(
                    ops,
                    center - tick,
                    open_y.min(close_y),
                    tick * 2.0,
                    (open_y - close_y).abs().max(1.0),
                    if rising {
                        CHART_BACKGROUND_COLOR
                    } else {
                        CHART_AXIS_COLOR
                    },
                );
                if rising {
                    push_line(
                        ops,
                        center - tick,
                        open_y.min(close_y),
                        center - tick,
                        open_y.max(close_y),
                        CHART_AXIS_COLOR,
                        1.0,
                    );
                }
            }
            Some(open) => {
                let open_y = scale.y(plot, value(open));
                push_line(
                    ops,
                    center - tick,
                    open_y,
                    center,
                    open_y,
                    CHART_AXIS_COLOR,
                    1.0,
                );
                let close_y = scale.y(plot, value(close));
                push_line(
                    ops,
                    center,
                    close_y,
                    center + tick,
                    close_y,
                    CHART_AXIS_COLOR,
                    1.0,
                );
            }
            None => {
                let close_y = scale.y(plot, value(close));
                push_line(
                    ops,
                    center,
                    close_y,
                    center + tick,
                    close_y,
                    CHART_AXIS_COLOR,
                    1.0,
                );
            }
        }
    }
}

/// Colour of the band `value` falls in, from a fixed contour ramp.
fn contour_color(ratio: f64) -> &'static str {
    const RAMP: [&str; 8] = [
        "#264478", "#4472C4", "#5B9BD5", "#70AD47", "#A9D18E", "#FFC000", "#ED7D31", "#9E480E",
    ];
    let slot = (ratio.clamp(0.0, 1.0) * (RAMP.len() - 1) as f64).round() as usize;
    RAMP[slot.min(RAMP.len() - 1)]
}

/// A 2D `c:surfaceChart` is Excel's contour chart: series are rows, categories
/// columns, and each cell takes the colour of its value band. A `c:wireframe`
/// surface draws the same mesh as lines. This is a flat contour, never a
/// projected 3D surface.
fn emit_surface<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    family: PlotFamily<'_>,
    plot: PlotArea,
) {
    let cat_count = category_count(family.series);
    let rows = family.series.len();
    if cat_count == 0 || rows == 0 {
        return;
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
    let scale = value_scale(family);
    let columns = cat_count.min(MAX_PLOT_SURFACE_CELLS / rows.max(1));
    let cell_w = plot.w / columns.max(1) as f64;
    let cell_h = plot.h / rows as f64;
    let wireframe = family.group.and_then(|group| group.wireframe) == Some(true);
    for (row, series) in family.series.iter().enumerate() {
        if ops.exhausted() {
            return;
        }
        let y = plot.y + plot.h - cell_h * (row + 1) as f64;
        for column in 0..columns {
            let x = plot.x + cell_w * category_position(family, column, columns) as f64;
            if wireframe {
                push_line(ops, x, y, x + cell_w, y, CHART_AXIS_COLOR, 0.5);
                push_line(ops, x, y, x, y + cell_h, CHART_AXIS_COLOR, 0.5);
            } else {
                let color = contour_color(scale.fraction(series.value(column)));
                push_rect(ops, x, y, cell_w, cell_h, color);
            }
        }
        push_text(
            ops,
            series.series.name.unwrap_or_default(),
            plot.x - 38.0,
            y + cell_h * 0.6,
            36.0,
            family.label(),
        );
    }
    for column in 0..columns {
        if ops.exhausted() {
            return;
        }
        push_text(
            ops,
            &category_label(family.series, column),
            plot.x + cell_w * category_position(family, column, columns) as f64 + 2.0,
            plot.y + plot.h + 14.0,
            cell_w - 4.0,
            family.label(),
        );
    }
}

fn push_path<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    plot: PlotArea,
    commands: Vec<GeometryPathCommand>,
    fill: &str,
    stroke: Option<PlotStroke>,
) {
    if commands.is_empty() {
        return;
    }
    ops.push(PlotOp::Path {
        x: plot.x,
        y: plot.y,
        w: plot.w,
        h: plot.h,
        commands,
        fill: fill.to_owned(),
        stroke,
    });
}

fn circle_path(cx: f64, cy: f64, r: f64) -> Vec<GeometryPathCommand> {
    polygon_path(cx, cy, r, 24, 0.0)
}

fn polygon_path(cx: f64, cy: f64, r: f64, sides: usize, rotation: f64) -> Vec<GeometryPathCommand> {
    let mut commands = Vec::with_capacity(sides + 1);
    for index in 0..sides {
        let angle = rotation + std::f64::consts::TAU * index as f64 / sides as f64;
        let (x, y) = (cx + r * angle.cos(), cy + r * angle.sin());
        commands.push(if index == 0 {
            GeometryPathCommand::Move { x, y }
        } else {
            GeometryPathCommand::Line { x, y }
        });
    }
    commands.push(GeometryPathCommand::Close);
    commands
}

/// The outline of one marker symbol, centred on `(cx, cy)` across `size`.
fn marker_path(symbol: PlotMarkerSymbol, cx: f64, cy: f64, size: f64) -> Vec<GeometryPathCommand> {
    let half = size / 2.0;
    let bar = |width: f64, height: f64| {
        vec![
            GeometryPathCommand::Move {
                x: cx - width,
                y: cy - height,
            },
            GeometryPathCommand::Line {
                x: cx + width,
                y: cy - height,
            },
            GeometryPathCommand::Line {
                x: cx + width,
                y: cy + height,
            },
            GeometryPathCommand::Line {
                x: cx - width,
                y: cy + height,
            },
            GeometryPathCommand::Close,
        ]
    };
    match symbol {
        PlotMarkerSymbol::Circle | PlotMarkerSymbol::Auto | PlotMarkerSymbol::None => {
            circle_path(cx, cy, half)
        }
        PlotMarkerSymbol::Dot => circle_path(cx, cy, half * 0.5),
        PlotMarkerSymbol::Diamond => polygon_path(cx, cy, half, 4, -std::f64::consts::FRAC_PI_2),
        PlotMarkerSymbol::Triangle => polygon_path(cx, cy, half, 3, -std::f64::consts::FRAC_PI_2),
        PlotMarkerSymbol::Square => polygon_path(cx, cy, half, 4, std::f64::consts::FRAC_PI_4),
        PlotMarkerSymbol::Star => star_path(cx, cy, half),
        PlotMarkerSymbol::Plus => cross_path(cx, cy, half, 0.0),
        PlotMarkerSymbol::X => cross_path(cx, cy, half, std::f64::consts::FRAC_PI_4),
        PlotMarkerSymbol::Dash => bar(half, half * 0.25),
    }
}

fn star_path(cx: f64, cy: f64, r: f64) -> Vec<GeometryPathCommand> {
    let mut commands = Vec::with_capacity(11);
    for index in 0..10 {
        let reach = if index % 2 == 0 { r } else { r * 0.42 };
        let angle = -std::f64::consts::FRAC_PI_2 + std::f64::consts::PI * index as f64 / 5.0;
        let (x, y) = (cx + reach * angle.cos(), cy + reach * angle.sin());
        commands.push(if index == 0 {
            GeometryPathCommand::Move { x, y }
        } else {
            GeometryPathCommand::Line { x, y }
        });
    }
    commands.push(GeometryPathCommand::Close);
    commands
}

/// A twelve-vertex cross, rotated by `rotation` to make a plus or an x.
fn cross_path(cx: f64, cy: f64, r: f64, rotation: f64) -> Vec<GeometryPathCommand> {
    const ARM: f64 = 0.3;
    let corners = [
        (-ARM, -1.0),
        (ARM, -1.0),
        (ARM, -ARM),
        (1.0, -ARM),
        (1.0, ARM),
        (ARM, ARM),
        (ARM, 1.0),
        (-ARM, 1.0),
        (-ARM, ARM),
        (-1.0, ARM),
        (-1.0, -ARM),
        (-ARM, -ARM),
    ];
    let (sin, cos) = rotation.sin_cos();
    let mut commands = Vec::with_capacity(corners.len() + 1);
    for (index, (dx, dy)) in corners.into_iter().enumerate() {
        let (x, y) = (
            cx + r * (dx * cos - dy * sin),
            cy + r * (dx * sin + dy * cos),
        );
        commands.push(if index == 0 {
            GeometryPathCommand::Move { x, y }
        } else {
            GeometryPathCommand::Line { x, y }
        });
    }
    commands.push(GeometryPathCommand::Close);
    commands
}

/// Draws `symbol` at `(x, y)`. A square stays a rectangle op, which is what
/// every host already paints fastest.
fn push_marker<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    symbol: Option<PlotMarkerSymbol>,
    x: f64,
    y: f64,
    size: f64,
    color: &str,
) {
    let Some(symbol) = symbol else {
        return;
    };
    if symbol == PlotMarkerSymbol::Square {
        push_rect(ops, x - size / 2.0, y - size / 2.0, size, size, color);
        return;
    }
    ops.push(PlotOp::Path {
        x: x - size / 2.0,
        y: y - size / 2.0,
        w: size,
        h: size,
        commands: marker_path(symbol, x, y, size),
        fill: color.to_owned(),
        stroke: None,
    });
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
    let available = ops.remaining;
    let total: f64 = (0..scanned)
        .map(|index| (index, series.value(index)))
        .filter(|(_, value)| *value > 0.0 && value.is_finite())
        .take(available)
        .map(|(_, value)| value)
        .sum();
    if total <= 0.0 {
        return;
    }
    let r = (width.min(height) * 0.34).max(10.0);
    let cx = x + width * 0.38;
    let cy = y + height * 0.46;
    let group = family.group;
    let inner_r = if family.chart_type == "doughnut" {
        r * group
            .and_then(|group| group.hole_size)
            .filter(|size| size.is_finite())
            .unwrap_or(50.0)
            .clamp(1.0, 90.0)
            / 100.0
    } else {
        0.0
    };
    let start = -std::f64::consts::FRAC_PI_2
        + group
            .and_then(|group| group.first_slice_angle)
            .filter(|angle| angle.is_finite())
            .unwrap_or(0.0)
            .rem_euclid(360.0)
            .to_radians();
    let vary = group.is_some_and(|group| group.vary_colors);
    let mut angle = start;
    for (index, value) in (0..scanned)
        .map(|index| (index, series.value(index)))
        .filter(|(_, value)| *value > 0.0 && value.is_finite())
        .take(available)
    {
        let sweep = (value / total) * std::f64::consts::TAU;
        let middle = angle + sweep / 2.0;
        let offset = r * series.explosion(index) / 100.0;
        let (ox, oy) = (cx + offset * middle.cos(), cy + offset * middle.sin());
        let color = if vary {
            series
                .point(index)
                .and_then(|point| point.color)
                .map(hex)
                .unwrap_or_else(|| {
                    CHART_SERIES_COLORS[index % CHART_SERIES_COLORS.len()].to_owned()
                })
        } else {
            series.point_color(index, index)
        };
        ops.push(PlotOp::Path {
            x: ox - r,
            y: oy - r,
            w: r * 2.0,
            h: r * 2.0,
            commands: pie_wedge_path(ox, oy, r, inner_r, angle, angle + sweep),
            fill: color,
            stroke: Some(PlotStroke {
                color: CHART_BACKGROUND_COLOR.to_owned(),
                width: 1.0,
            }),
        });
        let reach =
            r * pie_label_reach(point_label_spec(series, index).and_then(|labels| labels.position));
        push_point_label(
            ops,
            family,
            series,
            index,
            index,
            ox + reach * middle.cos(),
            oy + reach * middle.sin(),
            48.0,
            total,
        );
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

#[allow(clippy::too_many_arguments)]
fn emit_legend<S: PlotSink + ?Sized>(
    ops: &mut Emitter<'_, S>,
    chart: &PlotChart<'_>,
    budget: &mut ScanBudget,
    x: f64,
    y: f64,
    width: f64,
    style: &ResolvedText,
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
        push_text(ops, label, x + 12.0, yy + 8.0, width - 12.0, style);
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

/// A fraction as a percentage, snapped so a tick that lands on a whole
/// percent by arithmetic does not read as a fraction of one.
pub fn format_percent(value: f64) -> String {
    let scaled = value * 100.0;
    let snapped = if (scaled - scaled.round()).abs() < 1e-6 {
        scaled.round()
    } else {
        scaled
    };
    format_number(snapped) + "%"
}

/// Formats `value` through the subset of `c:numFmt` codes that changes what a
/// reader sees: decimal places, thousands grouping, a percent scale, and the
/// literal prefix and suffix around them. Anything richer — dates, fractions,
/// conditions, colours — returns `None` so the caller keeps its own format.
pub fn format_with_code(value: f64, code: &str) -> Option<String> {
    if !value.is_finite() || code.is_empty() {
        return None;
    }
    let section = code.split(';').next().unwrap_or(code);
    if section.eq_ignore_ascii_case("general") {
        return Some(format_number(value));
    }
    if section.contains(['y', 'd', 'h', 's', 'E', 'e', '?', '*', '[']) || section.contains("m/") {
        return None;
    }
    let digits = section
        .split('.')
        .nth(1)
        .map(|tail| {
            tail.chars()
                .take_while(|c| matches!(c, '0' | '#'))
                .count()
                .min(9)
        })
        .unwrap_or(0);
    let percent = section.contains('%');
    let scaled = if percent { value * 100.0 } else { value };
    let factor = 10_f64.powi(digits as i32);
    let rounded = (scaled.abs() * factor).round() / factor;
    let mut body = format!("{rounded:.digits$}");
    if section.contains(',') {
        body = group_thousands(&body);
    }
    let mut out = String::new();
    if scaled < 0.0 && !body.trim_start_matches(['0', '.', ',']).is_empty() {
        out.push('-');
    }
    out.push_str(&literal(section, true));
    out.push_str(&body);
    out.push_str(&literal(section, false));
    if percent {
        out.push('%');
    }
    Some(out)
}

/// The literal characters a format code puts before or after its digits.
fn literal(code: &str, leading: bool) -> String {
    let placeholder = |c: char| matches!(c, '0' | '#' | '.' | ',' | '%' | '?');
    let bytes: Vec<char> = code.chars().collect();
    let range: Box<dyn Iterator<Item = &char>> = if leading {
        Box::new(bytes.iter())
    } else {
        Box::new(bytes.iter().rev())
    };
    let mut literal: Vec<char> = range
        .take_while(|c| !placeholder(**c))
        .filter(|c| **c != '"' && **c != '\\' && **c != '_')
        .copied()
        .collect();
    if !leading {
        literal.reverse();
    }
    literal.into_iter().collect()
}

fn group_thousands(body: &str) -> String {
    let (whole, rest) = body.split_once('.').unwrap_or((body, ""));
    let mut grouped = String::with_capacity(whole.len() + whole.len() / 3 + rest.len() + 1);
    for (index, digit) in whole.chars().enumerate() {
        if index > 0 && (whole.len() - index) % 3 == 0 {
            grouped.push(',');
        }
        grouped.push(digit);
    }
    if !rest.is_empty() {
        grouped.push('.');
        grouped.push_str(rest);
    }
    grouped
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::{
        ChartDataLabels, ChartPlotGroup, ChartPointLabel, ChartSeries, ChartTextProperties,
    };

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

    fn labelled_space(
        chart_type: &str,
        group_labels: Option<ChartDataLabels>,
        series_labels: Option<ChartDataLabels>,
    ) -> ChartSpace {
        ChartSpace {
            chart_type: chart_type.to_owned(),
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some(chart_type.to_owned()),
                data_labels: group_labels,
                series: vec![ChartSeries {
                    categories: vec!["Q1".to_owned(), "Q2".to_owned()],
                    values: vec![3.0, 1.0],
                    color: "#4472C4".to_owned(),
                    data_labels: series_labels,
                    ..ChartSeries::default()
                }],
                ..ChartPlotGroup::default()
            }],
            ..ChartSpace::default()
        }
    }

    fn model_series(name: &str, values: &[f64]) -> ChartSeries {
        ChartSeries {
            name: Some(name.to_owned()),
            categories: values
                .iter()
                .enumerate()
                .map(|(index, _)| format!("Q{}", index + 1))
                .collect(),
            values: values.to_vec(),
            color: "#4472C4".to_owned(),
            ..ChartSeries::default()
        }
    }

    fn model_space(
        chart_type: &str,
        group_labels: Option<ChartDataLabels>,
        series: Vec<ChartSeries>,
    ) -> ChartSpace {
        ChartSpace {
            chart_type: chart_type.to_owned(),
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some(chart_type.to_owned()),
                data_labels: group_labels,
                series,
                ..ChartPlotGroup::default()
            }],
            ..ChartSpace::default()
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

    fn family<'a>(
        chart: &'a PlotChart<'a>,
        series: &'a [SeriesView<'a>],
        label: &'a ResolvedText,
    ) -> PlotFamily<'a> {
        PlotFamily {
            chart_type: chart.chart_type,
            series,
            value_axis: chart.value_axis,
            axis_titles: chart.axis_titles,
            group: chart.plot_groups.first(),
            chart_text: chart.text.chart,
            label,
            axis: None,
            x_axis: None,
            category_axis: None,
            secondary: false,
        }
    }

    #[test]
    fn fonts_render_the_css_shorthand() {
        assert_eq!(chart_label_font().css(), "400 10px Calibri, sans-serif");
        assert_eq!(chart_title_font().css(), "600 13px Calibri, sans-serif");
        assert_eq!(
            PlotTextStyle {
                font: Some("Georgia"),
                size_pt: Some(18.0),
                bold: Some(true),
                italic: Some(true),
                color: Some("FF0000"),
            }
            .resolve(CHART_LABEL_SIZE_PX, 400)
            .font
            .css(),
            "italic 700 24px Georgia"
        );
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
            if text == "Revenue" && *font == chart_title_font()));
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
        let label = PlotTextStyle::default().resolve(CHART_LABEL_SIZE_PX, 400);
        assert_eq!(value_range(family(&chart, &views, &label)), (-10.0, 10.0));
    }

    #[test]
    fn point_marker_without_a_size_falls_back_to_the_default() {
        let north = source(&[1.0, 2.0]);
        let mut series = series("North", &north);
        series.marker = Some(PlotMarker {
            size: Some(9.0),
            symbol: None,
        });
        series.points = vec![PlotPoint {
            index: Some(0),
            marker: Some(PlotMarker {
                size: None,
                symbol: None,
            }),
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
    fn the_label_budget_is_chart_wide_rather_than_per_series() {
        let values = vec![1.0; MAX_PLOT_OPS / 2 + 1];
        let wide = Source {
            categories: Vec::new(),
            values,
        };
        let labelled = |name, label| {
            let mut labelled = series(name, &wide);
            labelled.points = vec![PlotPoint {
                label: Some(label),
                ..PlotPoint::default()
            }];
            labelled
        };
        let chart = grouped(
            "line",
            PlotGroup {
                chart_type: Some("line"),
                series: vec![
                    labelled("First", "first-label"),
                    labelled("Second", "second-label"),
                ],
                markers: Some(false),
                ..PlotGroup::default()
            },
        );

        let ops = plot_chart(&chart, rect());

        assert_eq!(ops.len(), MAX_PLOT_OPS);
        let labels = texts(&ops);
        assert!(labels.contains(&"first-label".to_owned()));
        assert!(!labels.contains(&"second-label".to_owned()));
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
            plot_groups: vec![crate::chart::ChartPlotGroup {
                chart_type: Some("pie".to_owned()),
                series: vec![crate::chart::ChartSeries {
                    values: vec![1.0, 2.0],
                    color: "#4472C4".to_owned(),
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
                    ..crate::chart::ChartSeries::default()
                }],
                ..crate::chart::ChartPlotGroup::default()
            }],
            ..ChartSpace::default()
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

    fn rects(ops: &[PlotOp]) -> Vec<(f64, f64, f64, f64)> {
        ops.iter()
            .filter_map(|op| match op {
                PlotOp::Rect { x, y, w, h, .. } => Some((*x, *y, *w, *h)),
                _ => None,
            })
            .collect()
    }

    /// Every rectangle but the chart background and the legend swatches.
    fn bars(ops: &[PlotOp]) -> Vec<(f64, f64, f64, f64)> {
        rects(ops)
            .into_iter()
            .skip(1)
            .filter(|(_, _, w, h)| (*w - 8.0).abs() > 0.01 || (*h - 8.0).abs() > 0.01)
            .collect()
    }

    fn paths(ops: &[PlotOp]) -> usize {
        ops.iter()
            .filter(|op| matches!(op, PlotOp::Path { .. }))
            .count()
    }

    fn texts(ops: &[PlotOp]) -> Vec<String> {
        ops.iter()
            .filter_map(|op| match op {
                PlotOp::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect()
    }

    fn group<'a>(chart_type: &'a str, series: Vec<PlotSeries<'a>>) -> PlotGroup<'a> {
        PlotGroup {
            chart_type: Some(chart_type),
            series,
            ..PlotGroup::default()
        }
    }

    fn grouped<'a>(chart_type: &'a str, group: PlotGroup<'a>) -> PlotChart<'a> {
        PlotChart {
            chart_type,
            plot_groups: vec![group],
            ..PlotChart::default()
        }
    }

    #[test]
    fn every_family_draws_with_its_own_renderer() {
        let data = source(&[3.0, 1.0]);
        for chart_type in ["area", "radar"] {
            let chart = grouped(chart_type, group(chart_type, vec![series("North", &data)]));
            let ops = plot_chart(&chart, rect());
            let columns = bars(&ops)
                .into_iter()
                .filter(|(_, _, w, h)| *w > 8.0 && *h > 8.0)
                .count();
            assert_ne!(columns, 2, "{chart_type} still draws columns");
        }
        let area = grouped("area", group("area", vec![series("North", &data)]));
        assert_eq!(
            paths(&plot_chart(&area, rect())),
            1,
            "area fills one region"
        );
        let surface = grouped("surface", group("surface", vec![series("North", &data)]));
        assert!(
            plot_chart(&surface, rect())
                .iter()
                .any(|op| matches!(op, PlotOp::Rect { fill, .. } if fill == "#9E480E")),
            "a contour band takes its colour from the value ramp"
        );
    }

    #[test]
    fn scatter_places_points_at_their_x_values() {
        let categorical = source(&[10.0, 20.0]);
        let x = [0.0, 30.0];
        let mut xy = series("XY", &categorical);
        xy.x_values = &x;
        let chart = grouped("scatter", group("scatter", vec![xy]));
        let ops = plot_chart(&chart, rect());
        let markers: Vec<(f64, f64, f64, f64)> = rects(&ops)
            .into_iter()
            .filter(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
            .collect();
        assert_eq!(markers.len(), 2);
        assert!(markers[0].0 < markers[1].0);
        assert!(
            texts(&ops).contains(&"30".to_owned()),
            "the x axis labels its own values"
        );
    }

    #[test]
    fn scatter_and_bubble_never_substitute_indexes_for_missing_x_values() {
        let data = source(&[10.0, 20.0]);
        let x = [3.0];
        let sizes = [2.0, 4.0];
        let mut scatter_series = series("XY", &data);
        scatter_series.x_values = &x;
        let scatter = plot_chart(
            &grouped("scatter", group("scatter", vec![scatter_series])),
            rect(),
        );
        assert_eq!(
            rects(&scatter)
                .iter()
                .filter(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
                .count(),
            1
        );
        assert!(
            scatter
                .iter()
                .all(|op| !matches!(op, PlotOp::Line { width, .. } if *width == 2.0))
        );

        let mut bubble_series = series("Bubbles", &data);
        bubble_series.bubble_sizes = &sizes;
        let bubble = plot_chart(
            &grouped("bubble", group("bubble", vec![bubble_series])),
            rect(),
        );
        assert!(bubble.iter().all(|op| !matches!(op, PlotOp::Path { .. })));
        assert!(!texts(&bubble).contains(&"Q1".to_owned()));
        assert!(!texts(&bubble).contains(&"Q2".to_owned()));
    }

    #[test]
    fn scatter_and_bubble_skip_points_without_y_values() {
        let data = Source {
            categories: Vec::new(),
            values: vec![10.0],
        };
        let x = [1.0, 2.0];
        let sizes = [2.0, 4.0];
        let mut scatter_series = series("XY", &data);
        scatter_series.x_values = &x;
        let scatter = plot_chart(
            &grouped("scatter", group("scatter", vec![scatter_series])),
            rect(),
        );
        assert_eq!(
            rects(&scatter)
                .iter()
                .filter(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
                .count(),
            1
        );

        let mut bubble_series = series("Bubbles", &data);
        bubble_series.x_values = &x;
        bubble_series.bubble_sizes = &sizes;
        let bubble = plot_chart(
            &grouped("bubble", group("bubble", vec![bubble_series])),
            rect(),
        );
        assert_eq!(paths(&bubble), 1);
    }

    #[test]
    fn scatter_axis_references_are_independent_of_definition_order() {
        let data = Source {
            categories: Vec::new(),
            values: vec![80.0],
        };
        let x = [2.0];
        let mut xy = series("XY", &data);
        xy.x_values = &x;
        let mut scatter = group("scatter", vec![xy]);
        scatter.axis_ids = vec!["x", "y"];
        let x_axis = value_axis("x", 0.0, 10.0);
        let y_axis = value_axis("y", 0.0, 100.0);
        let chart = |axes| PlotChart {
            chart_type: "scatter",
            plot_groups: vec![scatter.clone()],
            axes,
            ..PlotChart::default()
        };

        assert_eq!(
            plot_chart(&chart(vec![x_axis, y_axis]), rect()),
            plot_chart(&chart(vec![y_axis, x_axis]), rect())
        );
    }

    #[test]
    fn unresolved_axis_references_never_select_another_definition() {
        let x_axis = value_axis("x", 0.0, 10.0);
        let y_axis = value_axis("y", 0.0, 100.0);
        let chart = PlotChart {
            chart_type: "scatter",
            axes: vec![x_axis, y_axis],
            ..PlotChart::default()
        };
        let resolve = |chart_type, axis_ids| {
            let group = PlotGroup {
                chart_type: Some(chart_type),
                axis_ids,
                ..PlotGroup::default()
            };
            let (x, y) = group_value_axes(&chart, &group);
            (x.and_then(|axis| axis.id), y.and_then(|axis| axis.id))
        };

        for chart_type in ["scatter", "bubble"] {
            assert_eq!(resolve(chart_type, Vec::new()), (None, None));
            assert_eq!(resolve(chart_type, vec!["missing", "y"]), (None, Some("y")));
            assert_eq!(resolve(chart_type, vec!["x"]), (Some("x"), None));
            assert_eq!(resolve(chart_type, vec!["x", "missing"]), (Some("x"), None));
            assert_eq!(resolve(chart_type, vec!["x", "x"]), (None, None));
        }

        let duplicated = PlotChart {
            chart_type: "scatter",
            axes: vec![x_axis, x_axis, y_axis],
            ..PlotChart::default()
        };
        let group = PlotGroup {
            chart_type: Some("scatter"),
            axis_ids: vec!["x", "y"],
            ..PlotGroup::default()
        };
        let (x, y) = group_value_axes(&duplicated, &group);
        assert!(x.is_none());
        assert_eq!(y.and_then(|axis| axis.id), Some("y"));
    }

    #[test]
    fn categorical_families_require_unique_referenced_axes() {
        let category = PlotAxis {
            id: Some("category"),
            kind: PlotAxisKind::Category,
            ..PlotAxis::default()
        };
        let value = value_axis("value", 0.0, 10.0);
        let chart = PlotChart {
            chart_type: "line",
            axes: vec![value, category],
            ..PlotChart::default()
        };
        for chart_type in ["bar", "column", "line", "area", "radar", "surface", "stock"] {
            let valid = PlotGroup {
                chart_type: Some(chart_type),
                axis_ids: vec!["category", "value"],
                ..PlotGroup::default()
            };
            assert_eq!(
                group_value_axes(&chart, &valid).1.and_then(|axis| axis.id),
                Some("value")
            );
            assert_eq!(
                group_category_axis(&chart, &valid).and_then(|axis| axis.id),
                Some("category")
            );
            for axis_ids in [Vec::new(), vec!["missing"]] {
                let unresolved = PlotGroup {
                    chart_type: Some(chart_type),
                    axis_ids,
                    ..PlotGroup::default()
                };
                assert!(group_value_axes(&chart, &unresolved).1.is_none());
                assert!(group_category_axis(&chart, &unresolved).is_none());
            }
        }

        let duplicated = PlotChart {
            chart_type: "line",
            axes: vec![value, value, category, category],
            ..PlotChart::default()
        };
        let group = PlotGroup {
            chart_type: Some("line"),
            axis_ids: vec!["category", "value"],
            ..PlotGroup::default()
        };
        assert!(group_value_axes(&duplicated, &group).1.is_none());
        assert!(group_category_axis(&duplicated, &group).is_none());
    }

    #[test]
    fn a_surface_category_axis_is_distinct_from_its_series_axis() {
        let category = PlotAxis {
            id: Some("category"),
            kind: PlotAxisKind::Category,
            ..PlotAxis::default()
        };
        let series = PlotAxis {
            id: Some("series"),
            kind: PlotAxisKind::Series,
            ..PlotAxis::default()
        };
        let chart = PlotChart {
            chart_type: "surface",
            axes: vec![series, value_axis("value", 0.0, 10.0), category],
            ..PlotChart::default()
        };
        let group = PlotGroup {
            chart_type: Some("surface"),
            axis_ids: vec!["category", "value", "series"],
            ..PlotGroup::default()
        };

        assert_eq!(
            group_category_axis(&chart, &group).and_then(|axis| axis.id),
            Some("category")
        );
    }

    #[test]
    fn a_scatter_style_without_lines_draws_only_markers() {
        let data = source(&[10.0, 20.0]);
        let x = [1.0, 2.0];
        let mut xy = series("XY", &data);
        xy.x_values = &x;
        let mut markers_only = group("scatter", vec![xy.clone()]);
        markers_only.scatter_style = Some("marker");
        let mut lines_only = group("scatter", vec![xy]);
        lines_only.scatter_style = Some("line");
        let series_lines = |chart: &PlotChart<'_>| {
            plot_chart(chart, rect())
                .iter()
                .filter(|op| matches!(op, PlotOp::Line { width, .. } if *width == 2.0))
                .count()
        };
        assert_eq!(series_lines(&grouped("scatter", markers_only)), 0);
        assert_eq!(series_lines(&grouped("scatter", lines_only)), 1);
    }

    #[test]
    fn bubbles_scale_by_area_or_by_width() {
        let data = source(&[10.0, 10.0]);
        let x = [1.0, 2.0];
        let sizes = [1.0, 4.0];
        let bubble = |represents: &'static str| {
            let mut series = series("Bubbles", &data);
            series.x_values = &x;
            series.bubble_sizes = &sizes;
            let mut group = group("bubble", vec![series]);
            group.size_represents = Some(represents);
            group
        };
        let radii = |chart: &PlotChart<'_>| {
            plot_chart(chart, rect())
                .iter()
                .filter_map(|op| match op {
                    PlotOp::Path { w, .. } => Some(*w),
                    _ => None,
                })
                .collect::<Vec<_>>()
        };
        let area = radii(&grouped("bubble", bubble("area")));
        let width = radii(&grouped("bubble", bubble("w")));
        assert_eq!(area.len(), 2);
        assert!((area[0] / area[1] - 0.5).abs() < 0.01, "{area:?}");
        assert!((width[0] / width[1] - 0.25).abs() < 0.01, "{width:?}");
    }

    #[test]
    fn a_bubble_without_coordinates_does_not_scale_visible_bubbles() {
        let data = source(&[10.0, 10.0]);
        let x = [1.0];
        let visible_size = [2.0];
        let hidden_size = [2.0, 400.0];
        let render = |sizes: &[f64]| {
            let mut bubbles = series("Bubbles", &data);
            bubbles.x_values = &x;
            bubbles.bubble_sizes = sizes;
            plot_chart(&grouped("bubble", group("bubble", vec![bubbles])), rect())
        };

        assert_eq!(render(&visible_size), render(&hidden_size));
    }

    #[test]
    fn a_stock_family_draws_hi_lo_and_up_down_bars() {
        let sources: Vec<Source> = [
            vec![10.0, 12.0],
            vec![20.0, 22.0],
            vec![5.0, 6.0],
            vec![18.0, 8.0],
        ]
        .into_iter()
        .map(|values| Source {
            categories: vec!["D1".to_owned(), "D2".to_owned()],
            values,
        })
        .collect();
        let names = ["Open", "High", "Low", "Close"];
        let ohlc: Vec<PlotSeries<'_>> = names
            .iter()
            .zip(&sources)
            .map(|(name, source)| series(name, source))
            .collect();
        let mut ohlc_group = group("stock", ohlc.clone());
        ohlc_group.hi_low_lines = true;
        ohlc_group.up_down_bars = true;
        let ops = plot_chart(&grouped("stock", ohlc_group), rect());
        let verticals = ops
            .iter()
            .filter(|op| {
                matches!(op, PlotOp::Line { x1, x2, y1, y2, .. }
                if (x1 - x2).abs() < 0.01 && (y1 - y2).abs() > 1.0 && (y1 - y2).abs() < 150.0)
            })
            .count();
        assert_eq!(verticals, 3, "two hi-lo lines and one rising open edge");
        let up_down: Vec<&str> = bars(&ops)
            .iter()
            .zip(
                ops.iter()
                    .filter_map(|op| match op {
                        PlotOp::Rect { fill, .. } => Some(fill.as_str()),
                        _ => None,
                    })
                    .skip(1),
            )
            .map(|(_, fill)| fill)
            .collect();
        assert!(
            up_down.contains(&CHART_AXIS_COLOR),
            "a falling bar fills dark"
        );
        assert!(
            up_down.contains(&CHART_BACKGROUND_COLOR),
            "a rising bar fills light"
        );
        let hlc = grouped("stock", group("stock", ohlc[1..].to_vec()));
        assert!(
            plot_chart(&hlc, rect())
                .iter()
                .any(|op| matches!(op, PlotOp::Line { .. })),
            "three series still plot high-low-close"
        );
    }

    #[test]
    fn stacked_bars_pile_onto_one_another() {
        let north = source(&[10.0, 20.0]);
        let south = source(&[5.0, 5.0]);
        let mut stacked = group(
            "column",
            vec![series("North", &north), series("South", &south)],
        );
        stacked.grouping = Some("stacked");
        let ops = plot_chart(&grouped("column", stacked), rect());
        let bars = bars(&ops);
        assert_eq!(bars.len(), 4);
        assert_eq!(bars[0].0, bars[1].0, "a stack shares one slot");
        assert!(bars[1].1 < bars[0].1, "the second series sits on the first");
        assert!(
            texts(&ops).contains(&"25".to_owned()),
            "the axis reaches the stack total"
        );
    }

    #[test]
    fn percent_stacked_normalizes_every_category() {
        let north = source(&[1.0, 30.0]);
        let south = source(&[3.0, 10.0]);
        let mut percent = group(
            "column",
            vec![series("North", &north), series("South", &south)],
        );
        percent.grouping = Some("percentStacked");
        let ops = plot_chart(&grouped("column", percent), rect());
        let bars = bars(&ops);
        assert_eq!(bars.len(), 4);
        let height = |slot: usize| bars[slot].3;
        assert!((height(0) + height(1) - (height(2) + height(3))).abs() < 0.01);
        assert!(texts(&ops).contains(&"100%".to_owned()));
    }

    #[test]
    fn gap_width_and_overlap_size_and_place_the_bars() {
        let north = source(&[10.0, 20.0]);
        let south = source(&[5.0, 5.0]);
        let sized = |gap: f64, overlap: f64| {
            let mut group = group(
                "column",
                vec![series("North", &north), series("South", &south)],
            );
            group.gap_width = Some(gap);
            group.overlap = Some(overlap);
            bars(&plot_chart(&grouped("column", group), rect()))
        };
        let wide = sized(0.0, 0.0);
        let narrow = sized(300.0, 0.0);
        assert!(wide[0].2 > narrow[0].2, "a wider gap leaves thinner bars");
        let apart = sized(150.0, 0.0);
        let over = sized(150.0, 100.0);
        assert!((apart[1].0 - apart[0].0 - apart[0].2).abs() < 0.01);
        assert!(
            (over[1].0 - over[0].0).abs() < 0.01,
            "full overlap shares a slot"
        );
    }

    #[test]
    fn marker_symbols_draw_their_own_outlines() {
        let data = source(&[1.0, 2.0]);
        for symbol in [
            PlotMarkerSymbol::Circle,
            PlotMarkerSymbol::Diamond,
            PlotMarkerSymbol::Triangle,
            PlotMarkerSymbol::Star,
            PlotMarkerSymbol::Plus,
            PlotMarkerSymbol::Dash,
            PlotMarkerSymbol::Dot,
            PlotMarkerSymbol::X,
            PlotMarkerSymbol::Auto,
        ] {
            let mut marked = series("North", &data);
            marked.marker = Some(PlotMarker {
                size: Some(8.0),
                symbol: Some(symbol),
            });
            let chart = grouped("line", group("line", vec![marked]));
            let ops = plot_chart(&chart, rect());
            assert_eq!(paths(&ops), 2, "{symbol:?} draws one outline per point");
        }
        let mut square = series("North", &data);
        square.marker = Some(PlotMarker {
            size: Some(8.0),
            symbol: Some(PlotMarkerSymbol::Square),
        });
        assert_eq!(
            paths(&plot_chart(
                &grouped("line", group("line", vec![square])),
                rect()
            )),
            0,
            "a square stays a rectangle op"
        );
        let mut none = series("North", &data);
        none.marker = Some(PlotMarker {
            size: Some(8.0),
            symbol: Some(PlotMarkerSymbol::None),
        });
        let bare = plot_chart(&grouped("line", group("line", vec![none])), rect());
        assert!(
            bars(&bare).is_empty(),
            "a none symbol draws nothing: {:?}",
            bars(&bare)
        );
    }

    #[test]
    fn a_line_group_marker_switch_removes_every_marker() {
        let data = source(&[1.0, 2.0]);
        let mut off = group("line", vec![series("North", &data)]);
        off.markers = Some(false);
        let ops = plot_chart(&grouped("line", off), rect());
        assert!(
            !rects(&ops)
                .iter()
                .any(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
        );
    }

    #[test]
    fn a_doughnut_honours_its_hole_size() {
        let data = source(&[3.0, 1.0]);
        let inner = |hole: Option<f64>| {
            let mut group = group("doughnut", vec![series("Share", &data)]);
            group.hole_size = hole;
            let ops = plot_chart(&grouped("doughnut", group), rect());
            let PlotOp::Path {
                commands,
                x,
                y,
                w,
                h,
                ..
            } = ops
                .iter()
                .find(|op| matches!(op, PlotOp::Path { .. }))
                .expect("a wedge")
            else {
                unreachable!()
            };
            let (cx, cy) = (x + w / 2.0, y + h / 2.0);
            commands
                .iter()
                .filter_map(|command| match command {
                    GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                        Some(((x - cx).powi(2) + (y - cy).powi(2)).sqrt())
                    }
                    _ => None,
                })
                .fold(f64::INFINITY, f64::min)
                / (w / 2.0)
        };
        assert!(inner(Some(20.0)) < inner(None));
        assert!(inner(None) < inner(Some(80.0)));
    }

    #[test]
    fn a_first_slice_angle_rotates_the_pie_and_explosion_offsets_a_slice() {
        let data = source(&[1.0, 1.0]);
        let start = |angle: Option<f64>| {
            let mut group = group("pie", vec![series("Share", &data)]);
            group.first_slice_angle = angle;
            let ops = plot_chart(&grouped("pie", group), rect());
            match ops.iter().find(|op| matches!(op, PlotOp::Path { .. })) {
                Some(PlotOp::Path { commands, .. }) => match commands[1] {
                    GeometryPathCommand::Line { x, y } => (x, y),
                    _ => unreachable!(),
                },
                _ => unreachable!(),
            }
        };
        let (upright_x, upright_y) = start(None);
        let (turned_x, turned_y) = start(Some(90.0));
        assert!((upright_x - turned_y).abs() > 1.0 || (upright_y - turned_x).abs() > 1.0);

        let mut exploded = series("Share", &data);
        exploded.points = vec![PlotPoint {
            index: Some(0),
            explosion: Some(50.0),
            ..PlotPoint::default()
        }];
        let ops = plot_chart(&grouped("pie", group("pie", vec![exploded])), rect());
        let centers: Vec<f64> = ops
            .iter()
            .filter_map(|op| match op {
                PlotOp::Path { x, w, .. } => Some(x + w / 2.0),
                _ => None,
            })
            .collect();
        assert!(
            (centers[0] - centers[1]).abs() > 1.0,
            "the exploded slice moves off centre"
        );
    }

    #[test]
    fn vary_colors_cycles_a_pie_through_the_palette() {
        let data = source(&[3.0, 1.0]);
        let mut varied = group("pie", vec![series("Share", &data)]);
        varied.vary_colors = true;
        let fills: Vec<String> = plot_chart(&grouped("pie", varied), rect())
            .iter()
            .filter_map(|op| match op {
                PlotOp::Path { fill, .. } => Some(fill.clone()),
                _ => None,
            })
            .collect();
        assert_eq!(fills, [CHART_SERIES_COLORS[0], CHART_SERIES_COLORS[1]]);
    }

    /// A chart with two value axes, the second one named by the second group.
    fn combo_with_axes<'a>(
        first: PlotGroup<'a>,
        second: PlotGroup<'a>,
        axes: Vec<PlotAxis<'a>>,
    ) -> PlotChart<'a> {
        PlotChart {
            chart_type: "column",
            plot_groups: vec![first, second],
            axes,
            ..PlotChart::default()
        }
    }

    fn value_axis<'a>(id: &'a str, min: f64, max: f64) -> PlotAxis<'a> {
        PlotAxis {
            id: Some(id),
            kind: PlotAxisKind::Value,
            range: PlotAxisRange {
                min: Some(min),
                max: Some(max),
            },
            major_gridlines: true,
            ..PlotAxis::default()
        }
    }

    #[test]
    fn a_secondary_value_axis_scales_only_the_group_that_names_it() {
        let small = source(&[10.0, 10.0]);
        let large = source(&[10.0, 10.0]);
        let mut primary = group("column", vec![series("Units", &small)]);
        primary.axis_ids = vec!["1", "2"];
        let mut secondary = group("line", vec![series("Rate", &large)]);
        secondary.axis_ids = vec!["1", "3"];
        let chart = combo_with_axes(
            primary,
            secondary,
            vec![value_axis("2", 0.0, 20.0), value_axis("3", 0.0, 100.0)],
        );
        let ops = plot_chart(&chart, rect());
        assert!(
            texts(&ops).contains(&"20".to_owned()) && texts(&ops).contains(&"100".to_owned()),
            "both axes label their own range"
        );
        let bar = bars(&ops)
            .into_iter()
            .find(|(_, _, w, h)| *w > 8.0 && *h > 8.0)
            .expect("a column");
        let marker = bars(&ops)
            .into_iter()
            .find(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
            .expect("a line marker");
        assert!(
            marker.1 > bar.1,
            "the same value sits lower against the wider axis"
        );
    }

    #[test]
    fn secondary_axis_groups_ignore_all_four_definition_positions() {
        let primary_data = source(&[10.0, 10.0]);
        let secondary_data = source(&[60.0, 60.0]);
        let mut primary = group("column", vec![series("Units", &primary_data)]);
        primary.axis_ids = vec!["cat-primary", "val-primary"];
        let mut secondary = group("line", vec![series("Rate", &secondary_data)]);
        secondary.axis_ids = vec!["cat-secondary", "val-secondary"];
        let category = |id| PlotAxis {
            id: Some(id),
            kind: PlotAxisKind::Category,
            ..PlotAxis::default()
        };
        let primary_category = category("cat-primary");
        let secondary_category = category("cat-secondary");
        let primary_value = value_axis("val-primary", 0.0, 20.0);
        let secondary_value = value_axis("val-secondary", 0.0, 100.0);
        let chart = |axes| combo_with_axes(primary.clone(), secondary.clone(), axes);

        assert_eq!(
            plot_chart(
                &chart(vec![
                    primary_category,
                    primary_value,
                    secondary_category,
                    secondary_value,
                ]),
                rect(),
            ),
            plot_chart(
                &chart(vec![
                    secondary_value,
                    primary_category,
                    primary_value,
                    secondary_category,
                ]),
                rect(),
            )
        );
    }

    #[test]
    fn a_log_axis_places_values_by_their_logarithm() {
        let data = source(&[1.0, 100.0]);
        let mut group = group("line", vec![series("Growth", &data)]);
        group.axis_ids = vec!["1"];
        let mut axis = value_axis("1", 1.0, 1000.0);
        axis.log_base = Some(10.0);
        let chart = PlotChart {
            chart_type: "line",
            plot_groups: vec![group],
            axes: vec![axis],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        let markers: Vec<(f64, f64, f64, f64)> = rects(&ops)
            .into_iter()
            .filter(|(_, _, w, h)| (*w - 4.0).abs() < 0.01 && (*h - 4.0).abs() < 0.01)
            .collect();
        let plot_h = 200.0 - 10.0 - 34.0;
        let travelled = markers[0].1 - markers[1].1;
        assert!(
            (travelled / plot_h - 2.0 / 3.0).abs() < 0.02,
            "two decades of three cover two thirds of the axis: {travelled}"
        );
    }

    #[test]
    fn a_reversed_axis_flips_the_value_direction_and_major_unit_places_the_ticks() {
        let data = source(&[0.0, 10.0]);
        let mut group = group("column", vec![series("North", &data)]);
        group.axis_ids = vec!["1"];
        let mut axis = value_axis("1", 0.0, 10.0);
        axis.reversed = true;
        axis.major_unit = Some(2.0);
        let chart = PlotChart {
            chart_type: "column",
            plot_groups: vec![group],
            axes: vec![axis],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        let labels = texts(&ops);
        for tick in ["0", "2", "4", "6", "8", "10"] {
            assert!(labels.contains(&tick.to_owned()), "{tick} is missing");
        }
        let tall = bars(&ops);
        assert_eq!(tall.len(), 2);
        assert!(
            tall[1].1 < 11.0 && tall[1].3 > 150.0,
            "a reversed axis grows the bar downward from the top: {tall:?}"
        );
    }

    #[test]
    fn gridlines_follow_the_axis_that_declares_them() {
        let data = source(&[1.0, 2.0]);
        let grid = |on: bool| {
            let mut group = group("column", vec![series("North", &data)]);
            group.axis_ids = vec!["1"];
            let mut axis = value_axis("1", 0.0, 4.0);
            axis.major_gridlines = on;
            let chart = PlotChart {
                chart_type: "column",
                plot_groups: vec![group],
                axes: vec![axis],
                ..PlotChart::default()
            };
            plot_chart(&chart, rect())
                .iter()
                .filter(|op| matches!(op, PlotOp::Line { color, .. } if color == CHART_GRID_COLOR))
                .count()
        };
        assert_eq!(grid(true), 5);
        assert_eq!(grid(false), 0);
    }

    #[test]
    fn a_hidden_axis_keeps_its_gridlines_but_drops_its_labels() {
        let data = source(&[1.0, 2.0]);
        let mut group = group("column", vec![series("North", &data)]);
        group.axis_ids = vec!["1"];
        let mut axis = value_axis("1", 0.0, 4.0);
        axis.hidden = true;
        let chart = PlotChart {
            chart_type: "column",
            plot_groups: vec![group],
            axes: vec![axis],
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        assert!(!texts(&ops).contains(&"4".to_owned()));
        assert_eq!(
            ops.iter()
                .filter(|op| matches!(op, PlotOp::Line { color, .. } if color == CHART_GRID_COLOR))
                .count(),
            5
        );
    }

    #[test]
    fn tick_marks_draw_only_when_the_axis_names_them() {
        let data = source(&[1.0, 2.0]);
        let ticks = |mark: Option<&'static str>| {
            let mut group = group("column", vec![series("North", &data)]);
            group.axis_ids = vec!["1"];
            let mut axis = value_axis("1", 0.0, 4.0);
            axis.major_tick_mark = mark;
            let chart = PlotChart {
                chart_type: "column",
                plot_groups: vec![group],
                axes: vec![axis],
                ..PlotChart::default()
            };
            plot_chart(&chart, rect())
                .iter()
                .filter(|op| {
                    matches!(op, PlotOp::Line { x1, x2, color, .. }
                    if color == CHART_AXIS_COLOR && (x1 - x2).abs() < 8.0 && (x1 - x2).abs() > 0.0)
                })
                .count()
        };
        assert_eq!(ticks(None), 0);
        assert_eq!(ticks(Some("none")), 0);
        assert_eq!(ticks(Some("out")), 5);
        assert_eq!(ticks(Some("cross")), 5);
    }

    #[test]
    fn a_reversed_category_axis_draws_the_categories_from_the_far_end() {
        let data = source(&[1.0, 9.0]);
        let mut group = group("column", vec![series("North", &data)]);
        group.axis_ids = vec!["1"];
        let chart = PlotChart {
            chart_type: "column",
            plot_groups: vec![group],
            axes: vec![PlotAxis {
                id: Some("1"),
                kind: PlotAxisKind::Category,
                reversed: true,
                ..PlotAxis::default()
            }],
            ..PlotChart::default()
        };
        let bars = bars(&plot_chart(&chart, rect()));
        assert_eq!(bars.len(), 2);
        assert!(
            bars[0].0 > bars[1].0,
            "the first category draws on the right"
        );
    }

    #[test]
    fn number_format_codes_change_what_the_axis_reads() {
        assert_eq!(format_with_code(0.256, "0.0%").as_deref(), Some("25.6%"));
        assert_eq!(format_with_code(1234.5, "#,##0").as_deref(), Some("1,235"));
        assert_eq!(
            format_with_code(1234.5, "\"$\"#,##0.00").as_deref(),
            Some("$1,234.50")
        );
        assert_eq!(
            format_with_code(-12.0, "0.0;(0.0)").as_deref(),
            Some("-12.0")
        );
        assert_eq!(format_with_code(7.0, "General").as_deref(), Some("7"));
        assert_eq!(format_with_code(7.0, "0 \"kg\"").as_deref(), Some("7 kg"));
        assert_eq!(format_with_code(7.0, "yyyy-mm-dd"), None);
        assert_eq!(format_with_code(f64::NAN, "0.0"), None);
        assert_eq!(format_percent(0.5), "50%");
    }

    #[test]
    fn an_axis_number_format_reaches_the_tick_labels() {
        let data = source(&[0.2, 0.8]);
        let mut group = group("column", vec![series("Share", &data)]);
        group.axis_ids = vec!["1"];
        let mut axis = value_axis("1", 0.0, 1.0);
        axis.number_format = Some("0%");
        let chart = PlotChart {
            chart_type: "column",
            plot_groups: vec![group],
            axes: vec![axis],
            ..PlotChart::default()
        };
        assert!(texts(&plot_chart(&chart, rect())).contains(&"75%".to_owned()));
    }

    #[test]
    fn data_label_switches_compose_the_text_they_ask_for() {
        let data = source(&[3.0, 1.0]);
        let compose = |spec: PlotDataLabels<'static>, chart_type: &'static str| {
            let mut labelled = series("North", &data);
            labelled.labels = Some(spec);
            texts(&plot_chart(
                &grouped(chart_type, group(chart_type, vec![labelled])),
                rect(),
            ))
        };
        let value = PlotDataLabels {
            show_value: true,
            ..PlotDataLabels::default()
        };
        assert!(compose(value, "column").contains(&"3".to_owned()));
        assert!(
            compose(
                PlotDataLabels {
                    show_category_name: true,
                    show_series_name: true,
                    separator: Some(" / "),
                    ..value
                },
                "column",
            )
            .contains(&"North / Q1 / 3".to_owned())
        );
        assert!(
            compose(
                PlotDataLabels {
                    show_percent: true,
                    show_value: false,
                    ..PlotDataLabels::default()
                },
                "pie",
            )
            .contains(&"75%".to_owned()),
            "a wedge takes its share of the series total"
        );
        assert!(
            compose(
                PlotDataLabels {
                    number_format: Some("0.00"),
                    ..value
                },
                "column",
            )
            .contains(&"3.00".to_owned())
        );
        let count = |texts: Vec<String>| texts.iter().filter(|text| *text == "3").count();
        assert!(
            count(compose(PlotDataLabels::default(), "column")) < count(compose(value, "column")),
            "a dLbls that shows nothing draws nothing"
        );
    }

    #[test]
    fn a_legend_key_draws_a_swatch_beside_its_label() {
        let data = source(&[3.0, 1.0]);
        let keyed = |show: bool| {
            let mut labelled = series("North", &data);
            labelled.color = Some("#123456");
            labelled.labels = Some(PlotDataLabels {
                show_value: true,
                show_legend_key: show,
                ..PlotDataLabels::default()
            });
            plot_chart(&grouped("column", group("column", vec![labelled])), rect())
                .iter()
                .filter(
                    |op| matches!(op, PlotOp::Rect { fill, w, .. } if fill == "#123456" && *w == 7.0),
                )
                .count()
        };
        assert_eq!(keyed(true), 2);
        assert_eq!(keyed(false), 0);
    }

    #[test]
    fn a_point_text_inherits_series_switches() {
        let space = labelled_space(
            "column",
            None,
            Some(ChartDataLabels {
                show_value: Some(true),
                show_legend_key: Some(true),
                points: Some(vec![ChartPointLabel {
                    index: Some(1.0),
                    text: Some("pinned".to_owned()),
                    labels: ChartDataLabels::default(),
                }]),
                ..ChartDataLabels::default()
            }),
        );
        let chart = PlotChart::from(&space);
        let point = &chart.plot_groups[0].series[0].points[0];
        let inherited = point.labels.expect("point labels");
        assert!(inherited.show_value);
        assert!(inherited.show_legend_key);
        let labels = texts(&plot_chart(&chart, rect()));
        assert!(labels.contains(&"pinned".to_owned()));
        assert!(labels.contains(&"3".to_owned()));
    }

    #[test]
    fn a_point_override_keeps_unset_series_switches() {
        let space = labelled_space(
            "column",
            None,
            Some(ChartDataLabels {
                show_value: Some(true),
                show_category_name: Some(true),
                points: Some(vec![ChartPointLabel {
                    index: Some(1.0),
                    text: None,
                    labels: ChartDataLabels {
                        show_value: Some(false),
                        ..ChartDataLabels::default()
                    },
                }]),
                ..ChartDataLabels::default()
            }),
        );
        let chart = PlotChart::from(&space);
        let labels = chart.plot_groups[0].series[0].points[0]
            .labels
            .expect("point labels");
        assert!(!labels.show_value);
        assert!(labels.show_category_name);
    }

    #[test]
    fn a_point_delete_suppresses_its_inherited_series_label() {
        let space = labelled_space(
            "pie",
            None,
            Some(ChartDataLabels {
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
            }),
        );
        let labels = texts(&plot_chart(&PlotChart::from(&space), rect()));
        assert!(labels.contains(&"3".to_owned()));
        assert!(!labels.contains(&"1".to_owned()));
    }

    #[test]
    fn a_series_delete_suppresses_group_labels() {
        let group = ChartDataLabels {
            show_value: Some(true),
            ..ChartDataLabels::default()
        };
        let series = ChartDataLabels {
            delete: Some(true),
            ..ChartDataLabels::default()
        };
        assert!(plot_labels_from_model(None, Some(&series), None, Some(&group)).is_none());
        let space = labelled_space("pie", Some(group), Some(series));
        let labels = texts(&plot_chart(&PlotChart::from(&space), rect()));
        assert!(!labels.contains(&"3".to_owned()));
        assert!(!labels.contains(&"1".to_owned()));
    }

    #[test]
    fn group_point_fields_resolve_through_the_four_level_cascade() {
        let group = ChartDataLabels {
            show_percent: Some(true),
            show_bubble_size: Some(true),
            position: Some("outEnd".to_owned()),
            text: Some(ChartTextProperties {
                font: Some("Group".to_owned()),
                color: Some("112233".to_owned()),
                ..ChartTextProperties::default()
            }),
            ..ChartDataLabels::default()
        };
        let group_point = ChartDataLabels {
            show_category_name: Some(true),
            show_legend_key: Some(true),
            separator: Some(" | ".to_owned()),
            text: Some(ChartTextProperties {
                italic: Some(true),
                ..ChartTextProperties::default()
            }),
            ..ChartDataLabels::default()
        };
        let series = ChartDataLabels {
            show_value: Some(false),
            position: Some("inEnd".to_owned()),
            text: Some(ChartTextProperties {
                size_pt: Some(12.0),
                bold: Some(false),
                ..ChartTextProperties::default()
            }),
            ..ChartDataLabels::default()
        };
        let series_point = ChartDataLabels {
            show_series_name: Some(true),
            number_format: Some("0.0".to_owned()),
            text: Some(ChartTextProperties {
                color: Some("AABBCC".to_owned()),
                ..ChartTextProperties::default()
            }),
            ..ChartDataLabels::default()
        };

        let labels = plot_labels_from_model(
            Some(&series_point),
            Some(&series),
            Some(&group_point),
            Some(&group),
        )
        .expect("labels resolve");

        assert!(!labels.show_value);
        assert!(labels.show_category_name);
        assert!(labels.show_series_name);
        assert!(labels.show_percent);
        assert!(labels.show_legend_key);
        assert!(labels.show_bubble_size);
        assert_eq!(labels.separator, Some(" | "));
        assert_eq!(labels.position, Some("inEnd"));
        assert_eq!(labels.number_format, Some("0.0"));
        assert_eq!(labels.text.font, Some("Group"));
        assert_eq!(labels.text.size_pt, Some(12.0));
        assert_eq!(labels.text.bold, Some(false));
        assert_eq!(labels.text.italic, Some(true));
        assert_eq!(labels.text.color, Some("AABBCC"));
    }

    #[test]
    fn series_delete_overrides_group_label_defaults() {
        let inherited = model_series("Inherited", &[41.0, 42.0]);
        let mut deleted = model_series("Deleted", &[51.0]);
        deleted.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            ..ChartDataLabels::default()
        });
        let space = model_space(
            "line",
            Some(ChartDataLabels {
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
            }),
            vec![inherited, deleted],
        );
        let chart = PlotChart::from(&space);
        let views = series_views(&chart.plot_groups[0].series, &mut ScanBudget::new());
        let label = PlotTextStyle::default().resolve(CHART_LABEL_SIZE_PX, 400);
        let family = family(&chart, &views, &label);

        assert_eq!(
            point_label(family, &views[0], 0, 0.0).as_deref(),
            Some("41")
        );
        assert!(point_label(family, &views[0], 1, 0.0).is_none());
        assert!(point_label(family, &views[1], 0, 0.0).is_none());
    }

    #[test]
    fn delete_is_inherited_until_a_lower_scope_overrides_it() {
        let mut inherited = model_series("Inherited", &[11.0, 12.0]);
        inherited.data_labels = Some(ChartDataLabels {
            show_category_name: Some(false),
            ..ChartDataLabels::default()
        });
        let mut point_restored = model_series("Point", &[21.0, 22.0]);
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
        let mut series_restored = model_series("Series", &[31.0, 32.0]);
        series_restored.data_labels = Some(ChartDataLabels {
            delete: Some(false),
            ..ChartDataLabels::default()
        });
        let space = model_space(
            "line",
            Some(ChartDataLabels {
                delete: Some(true),
                show_value: Some(true),
                ..ChartDataLabels::default()
            }),
            vec![inherited, point_restored, series_restored],
        );
        let chart = PlotChart::from(&space);
        let views = series_views(&chart.plot_groups[0].series, &mut ScanBudget::new());
        let label = PlotTextStyle::default().resolve(CHART_LABEL_SIZE_PX, 400);
        let family = family(&chart, &views, &label);

        assert!(point_label(family, &views[0], 0, 0.0).is_none());
        assert!(point_label(family, &views[0], 1, 0.0).is_none());
        assert!(point_label(family, &views[1], 0, 0.0).is_none());
        assert_eq!(
            point_label(family, &views[1], 1, 0.0).as_deref(),
            Some("22")
        );
        assert_eq!(
            point_label(family, &views[2], 0, 0.0).as_deref(),
            Some("31")
        );
        assert_eq!(
            point_label(family, &views[2], 1, 0.0).as_deref(),
            Some("32")
        );
    }

    #[test]
    fn series_and_point_label_overrides_do_not_leak() {
        let mut shown = model_series("Shown", &[11.0, 12.0]);
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
        let inherited = model_series("Inherited", &[21.0, 22.0]);
        let mut deleted = model_series("Deleted", &[31.0, 32.0]);
        deleted.data_labels = Some(ChartDataLabels {
            delete: Some(true),
            ..ChartDataLabels::default()
        });
        let space = model_space("line", None, vec![shown, inherited, deleted]);
        let chart = PlotChart::from(&space);
        let views = series_views(&chart.plot_groups[0].series, &mut ScanBudget::new());
        let label = PlotTextStyle::default().resolve(CHART_LABEL_SIZE_PX, 400);
        let family = family(&chart, &views, &label);

        assert_eq!(
            point_label(family, &views[0], 0, 0.0).as_deref(),
            Some("11")
        );
        assert!(point_label(family, &views[0], 1, 0.0).is_none());
        for view in &views[1..] {
            assert!(point_label(family, view, 0, 0.0).is_none());
            assert!(point_label(family, view, 1, 0.0).is_none());
        }
    }

    #[test]
    fn a_label_position_moves_the_text_within_its_bar() {
        let data = source(&[10.0, 20.0]);
        let placed = |position: &'static str| {
            let mut labelled = series("North", &data);
            labelled.labels = Some(PlotDataLabels {
                show_value: true,
                position: Some(position),
                ..PlotDataLabels::default()
            });
            plot_chart(&grouped("column", group("column", vec![labelled])), rect())
                .into_iter()
                .filter_map(|op| match op {
                    PlotOp::Text {
                        text, baseline_y, ..
                    } if text == "20" => Some(baseline_y),
                    _ => None,
                })
                .next_back()
                .expect("a label")
        };
        assert!(placed("outEnd") < placed("inEnd"));
        assert!(placed("inEnd") < placed("ctr"));
        assert!(placed("ctr") < placed("inBase"));
    }

    #[test]
    fn chart_text_properties_reach_every_scope_and_inherit() {
        let data = source(&[1.0, 2.0]);
        let mut labelled = series("North", &data);
        labelled.labels = Some(PlotDataLabels {
            show_value: true,
            number_format: Some("0.0"),
            text: PlotTextStyle {
                size_pt: Some(6.0),
                ..PlotTextStyle::default()
            },
            ..PlotDataLabels::default()
        });
        let axis = PlotAxis {
            id: Some("1"),
            kind: PlotAxisKind::Value,
            range: PlotAxisRange {
                min: Some(0.0),
                max: Some(4.0),
            },
            major_gridlines: true,
            text: PlotTextStyle {
                italic: Some(true),
                ..PlotTextStyle::default()
            },
            ..PlotAxis::default()
        };
        let mut group = group("column", vec![labelled]);
        group.axis_ids = vec!["1"];
        let chart = PlotChart {
            chart_type: "column",
            title: Some("Revenue"),
            plot_groups: vec![group],
            axes: vec![axis],
            text: PlotChartText {
                chart: PlotTextStyle {
                    font: Some("Georgia"),
                    color: Some("#112233"),
                    ..PlotTextStyle::default()
                },
                title: PlotTextStyle {
                    size_pt: Some(30.0),
                    ..PlotTextStyle::default()
                },
                legend: PlotTextStyle {
                    bold: Some(true),
                    ..PlotTextStyle::default()
                },
            },
            ..PlotChart::default()
        };
        let ops = plot_chart(&chart, rect());
        let font = |needle: &str| {
            ops.iter()
                .find_map(|op| match op {
                    PlotOp::Text {
                        text, font, color, ..
                    } if text == needle => Some((font.css(), color.clone())),
                    _ => None,
                })
                .unwrap_or_else(|| panic!("{needle} is missing"))
        };
        assert_eq!(
            font("Revenue"),
            ("600 40px Georgia".to_owned(), "#112233".to_owned())
        );
        assert_eq!(font("4").0, "italic 400 10px Georgia");
        assert_eq!(font("North").0, "700 10px Georgia");
        assert_eq!(font("2.0").0, "400 8px Georgia");
    }

    #[test]
    fn every_family_stays_finite_and_bounded_on_hostile_input() {
        let extremes = Source {
            categories: Vec::new(),
            values: vec![f64::MAX, -f64::MAX, 5.0, f64::NAN],
        };
        let wide: Vec<f64> = (0..50_000).map(|value| value as f64).collect();
        let dense = Source {
            categories: Vec::new(),
            values: wide.clone(),
        };
        for chart_type in [
            "area", "scatter", "bubble", "radar", "stock", "surface", "doughnut", "line", "column",
            "bar",
        ] {
            for data in [&extremes, &dense] {
                let mut plotted = series("Hostile", data);
                plotted.x_values = &wide;
                plotted.bubble_sizes = &wide;
                let chart = grouped(chart_type, group(chart_type, vec![plotted]));
                let ops = plot_chart(&chart, rect());
                assert!(ops.len() <= MAX_PLOT_OPS, "{chart_type}");
                for op in &ops {
                    assert!(op_is_finite(op), "{chart_type} {op:?}");
                }
            }
        }
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
            fn push_op(&mut self, _: PlotOp) -> bool {
                self.0 += 1;
                true
            }
        }
        let mut counter = Counter(0);
        plot_chart_into(&chart, rect(), &mut counter);
        assert_eq!(counter.0, plot_chart(&chart, rect()).len());
    }
}
