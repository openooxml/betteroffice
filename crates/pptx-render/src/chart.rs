//! Chart plotting: shared plot ops translated into slide primitives.

use ooxml_drawingml::GeometryPathCommand;
use ooxml_drawingml::chart::{
    ChartSpace, PlotChart, PlotDataLabels, PlotFont, PlotOp, PlotRect, PlotSink, chart_aria_label,
    plot_chart_into,
};

use crate::{Paint, Primitive, RenderError, Stroke, Transform};

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
    let chart = plot_model(space);
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

/// The shared geometry reads `c:dLbls` itself. A part that only carries the
/// legacy `show_data_labels` flag — one with no switches of its own — still
/// gets the values Excel would default to.
fn plot_model(space: &ChartSpace) -> PlotChart<'_> {
    let mut chart = PlotChart::from(space);
    for (group, plotted) in space.plot_groups.iter().zip(chart.plot_groups.iter_mut()) {
        if !group.show_data_labels {
            continue;
        }
        for (model, series) in group.series.iter().zip(plotted.series.iter_mut()) {
            let declared = model.data_labels.is_some() || group.data_labels.is_some();
            if declared && series.labels.is_none() {
                continue;
            }
            if series.labels.is_none_or(|labels| !labels.shows_anything()) {
                series.labels = Some(PlotDataLabels {
                    show_value: true,
                    ..PlotDataLabels::default()
                });
            }
        }
    }
    chart
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
    fn accepts_more(&mut self) -> bool {
        self.remaining > 0 && self.error.is_none()
    }

    fn push_op(&mut self, op: PlotOp) -> bool {
        if self.remaining == 0 || self.error.is_some() {
            return false;
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
                        return false;
                    }
                }
            }
        };
        self.primitives.push(primitive);
        true
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
    use pptx_parse::{
        ChartAxes, ChartAxis, ChartDataLabels, ChartLegend, ChartMarker, ChartPlotGroup,
        ChartPoint, ChartSeries,
    };

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
            ..ChartSeries::default()
        }
    }

    fn group(chart_type: &str, series: Vec<ChartSeries>) -> ChartPlotGroup {
        ChartPlotGroup {
            chart_type: Some(chart_type.to_owned()),
            grouping: Some("clustered".to_owned()),
            series,
            ..ChartPlotGroup::default()
        }
    }

    fn space(chart_type: &str, groups: Vec<ChartPlotGroup>) -> ChartSpace {
        ChartSpace {
            chart_type: chart_type.to_owned(),
            title: Some("Revenue".to_owned()),
            legend: Some(ChartLegend {
                position: Some("right".to_owned()),
                visible: true,
                text: None,
            }),
            series: groups
                .iter()
                .flat_map(|group| group.series.iter())
                .cloned()
                .collect(),
            plot_groups: groups,
            ..ChartSpace::default()
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
                text: None,
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
            text: None,
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
                title: Some("Millions".to_owned()),
                min: Some(0.0),
                max: Some(10.0),
                axis_type: "value".to_owned(),
                position: Some("left".to_owned()),
                ..ChartAxis::default()
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
            color: None,
        });
        point_series.points = Some(vec![ChartPoint {
            index: Some(1.0),
            explosion: None,
            color: "#AABBCC".to_owned(),
        }]);
        let mut group = group("line", vec![point_series]);
        group.show_data_labels = true;
        let chart = plot(&space("line", vec![group]));

        assert!(fills(&chart).contains(&"#AABBCC"));
        assert!(
            parts(&chart).iter().any(
                |primitive| matches!(primitive, Primitive::Shape { geometry, w, .. }
                if geometry == "custom" && (*w - 11.0).abs() < 0.001)
            ),
            "a circle symbol draws its own outline at the marker size"
        );
        for value in ["3", "1", "2"] {
            assert!(texts(&chart).contains(&value.to_owned()), "{value}");
        }
    }

    #[test]
    fn the_part_budget_still_bounds_a_chart_that_labels_every_point() {
        let values = (0..3_000).map(|value| value as f64).collect::<Vec<_>>();
        let mut group = group(
            "line",
            vec![
                series("First", &values, "#112233"),
                series("Second", &values, "#445566"),
            ],
        );
        group.show_data_labels = true;
        let chart = chart_primitive(
            frame("Wide"),
            &space("line", vec![group]),
            512,
            &mut |text| {
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
            },
        )
        .expect("chart plots");
        assert_eq!(parts(&chart).len(), 512);
    }

    #[test]
    fn a_series_that_switches_its_own_labels_off_draws_none() {
        let labelled = |delete: Option<bool>| {
            let mut labelled_series = series("North", &[3.0, 1.0], "#112233");
            labelled_series.data_labels = delete.map(|delete| ChartDataLabels {
                delete: Some(delete),
                show_value: Some(true),
                ..ChartDataLabels::default()
            });
            let mut group = group("line", vec![labelled_series]);
            group.show_data_labels = true;
            texts(&plot(&space("line", vec![group]))).len()
        };
        assert!(labelled(Some(true)) < labelled(Some(false)));
        assert_eq!(labelled(None), labelled(Some(false)));
    }

    #[test]
    fn label_switches_compose_the_text_the_part_asks_for() {
        let mut labelled = series("North", &[3.0, 1.0], "#112233");
        labelled.data_labels = Some(ChartDataLabels {
            show_value: Some(true),
            show_category_name: Some(true),
            show_percent: Some(true),
            separator: Some(" | ".to_owned()),
            ..ChartDataLabels::default()
        });
        let mut group = group("pie", vec![labelled]);
        group.show_data_labels = true;
        assert!(texts(&plot(&space("pie", vec![group]))).contains(&"Q1 | 3 | 75%".to_owned()));
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
        group.show_data_labels = true;
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
