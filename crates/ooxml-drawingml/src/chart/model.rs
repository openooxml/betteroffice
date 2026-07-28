use serde::{Deserialize, Serialize};

/// A parsed `c:chartSpace`, free of any host-package metadata.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSpace {
    pub chart_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend: Option<ChartLegend>,
    pub series: Vec<ChartSeries>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axes: Option<ChartAxes>,
    pub plot_groups: Vec<ChartPlotGroup>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_list: Option<Vec<ChartAxis>>,
    /// `c:txPr` on `c:chartSpace`: the root of chart text inheritance.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ChartTextProperties>,
    /// `c:txPr` on `c:title`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_text: Option<ChartTextProperties>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ChartLegend {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    pub visible: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ChartTextProperties>,
}

/// Run properties read off a `c:txPr`, all optional so an absent one inherits.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextProperties {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_pt: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

impl ChartTextProperties {
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
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
    /// `c:xVal` of a scatter or bubble series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_values: Option<Vec<f64>>,
    /// `c:bubbleSize` of a bubble series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_sizes: Option<Vec<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<ChartDataLabels>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ChartMarker {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub symbol: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ChartPoint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub explosion: Option<f64>,
    pub color: String,
}

/// A `c:dLbls`, every switch optional so an unset one inherits its parent.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataLabels {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_value: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_category_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_series_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ChartTextProperties>,
    /// Per-point `c:dLbl` overrides; never nested further.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub points: Option<Vec<ChartPointLabel>>,
}

/// One `c:dLbl`: an index plus the switches that override the series default.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartPointLabel {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub index: Option<f64>,
    /// Literal `c:tx` text, which replaces every composed field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    pub labels: ChartDataLabels,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
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
    /// `c:scatterStyle`: `lineMarker`, `marker`, `smoothMarker`, `line`, `smooth`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    /// `c:radarStyle`: `standard`, `marker`, `filled`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    /// `c:bubbleScale`, a percentage of the default bubble size.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_scale: Option<f64>,
    /// `c:sizeRepresents`: `area` or `w`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_represents: Option<String>,
    /// `c:wireframe` of a surface chart.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<bool>,
    /// `c:hiLowLines` of a stock or line chart.
    #[serde(default)]
    pub hi_low_lines: bool,
    /// `c:upDownBars` of a stock or line chart.
    #[serde(default)]
    pub up_down_bars: bool,
    /// `c:marker` of a line chart, which switches every series marker off.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<ChartDataLabels>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ChartAxes {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<ChartAxis>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<ChartAxis>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
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
    #[serde(default)]
    pub major_gridlines: bool,
    #[serde(default)]
    pub minor_gridlines: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ChartTextProperties>,
}
