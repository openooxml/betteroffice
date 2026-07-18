use serde::{Deserialize, Serialize};

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
