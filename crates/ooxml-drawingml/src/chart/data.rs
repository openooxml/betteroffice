use serde::{Deserialize, Serialize};

use super::geometry::{
    PlotChart, PlotSeries, PlottedSeries, format_number, format_value, plotted_series,
};

/// The most cells (series times categories) a [`ChartData`] holds.
pub const MAX_CHART_DATA_CELLS: usize = 10_000;
/// The most text, in bytes, a [`ChartData`] holds across its labels and values.
pub const MAX_CHART_DATA_TEXT_BYTES: usize = 1024 * 1024;

/// What a chart plots, as text for assistive technology: one row per category
/// and one column per series.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartData {
    pub categories: Vec<String>,
    pub series: Vec<ChartDataSeries>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataSeries {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Each category's value in the number format its data label uses, else
    /// the format the values were cached with; `None` where the cache has no
    /// number.
    pub values: Vec<Option<String>>,
}

/// The chart's data, or `None` when it plots nothing or would exceed
/// [`MAX_CHART_DATA_CELLS`] or [`MAX_CHART_DATA_TEXT_BYTES`].
pub fn chart_data(chart: &PlotChart<'_>) -> Option<ChartData> {
    let plotted = plotted_series(chart);
    let series: Vec<&PlotSeries<'_>> = plotted.iter().map(PlottedSeries::series).collect();
    let count = series
        .iter()
        .map(|series| {
            series
                .categories
                .len()
                .max(series.values.len())
                .max(series.x_values.len())
        })
        .max()
        .unwrap_or(0);
    if series.is_empty() || count == 0 || series.len().checked_mul(count)? > MAX_CHART_DATA_CELLS {
        return None;
    }
    let mut bytes = 0usize;
    let mut charge = |text: &str| -> Option<()> {
        bytes = bytes.checked_add(text.len())?;
        (bytes <= MAX_CHART_DATA_TEXT_BYTES).then_some(())
    };
    let mut categories = Vec::with_capacity(count);
    for index in 0..count {
        let label = series
            .iter()
            .find_map(|series| series.categories.get(index).cloned())
            .or_else(|| {
                series
                    .iter()
                    .find_map(|series| series.x_values.get(index))
                    .filter(|value| value.is_finite())
                    .map(|value| format_number(*value))
            })
            .unwrap_or_else(|| (index + 1).to_string());
        charge(&label)?;
        categories.push(label);
    }
    let mut columns = Vec::with_capacity(series.len());
    for (plotted, series) in plotted.iter().zip(series) {
        if let Some(name) = series.name {
            charge(name)?;
        }
        let mut values = Vec::with_capacity(count);
        for index in 0..count {
            let value = series
                .values
                .get(index)
                .filter(|value| value.is_finite())
                .map(|value| {
                    let code = plotted.label_number_format(index).or(series.value_format);
                    format_value(*value, code)
                });
            if let Some(text) = &value {
                charge(text)?;
            }
            values.push(value);
        }
        columns.push(ChartDataSeries {
            name: series.name.map(str::to_owned),
            values,
        });
    }
    Some(ChartData {
        categories,
        series: columns,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::{PlotDataLabels, PlotGroup, PlotPoint};

    fn series<'a>(name: &'a str, categories: &'a [String], values: &'a [f64]) -> PlotSeries<'a> {
        PlotSeries {
            name: Some(name),
            categories,
            values,
            ..PlotSeries::default()
        }
    }

    fn texts(values: &[Option<String>]) -> Vec<Option<&str>> {
        values.iter().map(Option::as_deref).collect()
    }

    #[test]
    fn the_series_read_through_their_cached_format() {
        let categories = vec!["Q1".to_owned(), "Q2".to_owned()];
        let values = [0.86, f64::NAN];
        let chart = PlotChart {
            series: vec![PlotSeries {
                value_format: Some("0%"),
                ..series("Share", &categories, &values)
            }],
            ..PlotChart::default()
        };
        let data = chart_data(&chart).unwrap();
        assert_eq!(data.categories, categories);
        assert_eq!(data.series[0].name.as_deref(), Some("Share"));
        assert_eq!(texts(&data.series[0].values), [Some("86%"), None]);
    }

    #[test]
    fn plot_groups_win_over_the_flat_series_as_the_drawing_does() {
        let categories = vec!["Q1".to_owned()];
        let values = [0.5];
        let chart = PlotChart {
            series: vec![series("Flat", &categories, &values)],
            plot_groups: vec![PlotGroup {
                series: vec![PlotSeries {
                    value_format: Some("0%"),
                    ..series("Grouped", &categories, &values)
                }],
                ..PlotGroup::default()
            }],
            ..PlotChart::default()
        };
        let data = chart_data(&chart).unwrap();
        assert_eq!(data.series.len(), 1);
        assert_eq!(data.series[0].name.as_deref(), Some("Grouped"));
        assert_eq!(texts(&data.series[0].values), [Some("50%")]);
    }

    #[test]
    fn a_label_number_format_wins_over_the_cache_and_a_point_over_its_series() {
        let categories = vec!["A".to_owned(), "B".to_owned()];
        let values = [12.0, 7.0];
        let labels = |code| PlotDataLabels {
            number_format: Some(code),
            ..PlotDataLabels::default()
        };
        let chart = PlotChart {
            series: vec![PlotSeries {
                value_format: Some("0"),
                labels: Some(labels("0.00")),
                points: vec![PlotPoint {
                    index: Some(1),
                    labels: Some(labels("0.0")),
                    ..PlotPoint::default()
                }],
                ..series("S", &categories, &values)
            }],
            ..PlotChart::default()
        };
        let data = chart_data(&chart).unwrap();
        assert_eq!(texts(&data.series[0].values), [Some("12.00"), Some("7.0")]);
    }

    #[test]
    fn point_labels_past_the_drawing_scan_budget_are_read_as_the_drawing_reads_them() {
        let first = [1.0];
        let later = [5.0];
        let labels = |code| PlotDataLabels {
            number_format: Some(code),
            ..PlotDataLabels::default()
        };
        let spent: Vec<PlotPoint<'_>> = (0..crate::chart::MAX_PLOT_DATA_SCAN)
            .map(|index| PlotPoint {
                index: Some(index),
                ..PlotPoint::default()
            })
            .collect();
        let chart = PlotChart {
            series: vec![
                PlotSeries {
                    points: spent,
                    ..series("Spends the scan", &[], &first)
                },
                PlotSeries {
                    labels: Some(labels("0.00")),
                    points: vec![PlotPoint {
                        index: Some(0),
                        labels: Some(labels("0.0")),
                        ..PlotPoint::default()
                    }],
                    ..series("Later", &[], &later)
                },
            ],
            ..PlotChart::default()
        };
        let data = chart_data(&chart).unwrap();
        assert_eq!(
            texts(&data.series[1].values),
            [Some("5.00")],
            "the drawing scans no points of the later series, so its series label format applies"
        );
    }

    #[test]
    fn a_combo_chart_reads_every_group_and_pads_short_series() {
        let categories = vec!["A".to_owned(), "B".to_owned(), "C".to_owned()];
        let (bars, line) = ([1.0, 2.0, 3.0], [4.5]);
        let chart = PlotChart {
            plot_groups: vec![
                PlotGroup {
                    series: vec![series("Bars", &categories, &bars)],
                    ..PlotGroup::default()
                },
                PlotGroup {
                    series: vec![series("Line", &[], &line)],
                    ..PlotGroup::default()
                },
            ],
            ..PlotChart::default()
        };
        let data = chart_data(&chart).unwrap();
        assert_eq!(data.categories, categories);
        assert_eq!(texts(&data.series[1].values), [Some("4.5"), None, None]);
    }

    #[test]
    fn scatter_x_values_name_the_rows_and_numbers_stand_in_otherwise() {
        let (x, y) = ([0.5, 1.5], [10.0, 20.0]);
        let scatter = PlotChart {
            series: vec![PlotSeries {
                x_values: &x,
                ..series("Y", &[], &y)
            }],
            ..PlotChart::default()
        };
        assert_eq!(chart_data(&scatter).unwrap().categories, ["0.5", "1.5"]);
        let plain = PlotChart {
            series: vec![series("Y", &[], &y)],
            ..PlotChart::default()
        };
        assert_eq!(chart_data(&plain).unwrap().categories, ["1", "2"]);
    }

    #[test]
    fn an_empty_chart_has_no_data() {
        assert_eq!(chart_data(&PlotChart::default()), None);
        let empty = PlotChart {
            series: vec![series("None", &[], &[])],
            ..PlotChart::default()
        };
        assert_eq!(chart_data(&empty), None);
    }

    #[test]
    fn data_past_the_cell_cap_is_left_out() {
        let values = vec![1.0; MAX_CHART_DATA_CELLS / 2 + 1];
        let within = PlotChart {
            series: vec![series("One", &[], &values)],
            ..PlotChart::default()
        };
        assert!(chart_data(&within).is_some());
        let beyond = PlotChart {
            series: vec![series("One", &[], &values), series("Two", &[], &values)],
            ..PlotChart::default()
        };
        assert_eq!(chart_data(&beyond), None);
    }

    #[test]
    fn data_past_the_text_cap_is_left_out() {
        let long = "x".repeat(MAX_CHART_DATA_TEXT_BYTES / 2 + 1);
        let fits = vec![long.clone()];
        let values = [1.0, 2.0];
        let within = PlotChart {
            series: vec![series("S", &fits, &values)],
            ..PlotChart::default()
        };
        assert!(chart_data(&within).is_some());
        let over = vec![long.clone(), long];
        let beyond = PlotChart {
            series: vec![series("S", &over, &values)],
            ..PlotChart::default()
        };
        assert_eq!(chart_data(&beyond), None);
    }
}
