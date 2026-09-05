use std::collections::HashMap;

use crate::GeometryPathCommand;

const ELLIPSE_KAPPA: f64 = 0.552_284_749_830_793_6;
const ROUND_RECT_ADJUSTMENT: f64 = 0.166_67;

pub fn preset_geometry_default_adjustments(shape_type: &str) -> HashMap<String, f64> {
    let values = match shape_type {
        "roundRect" => vec![("adj", ROUND_RECT_ADJUSTMENT)],
        "triangle" | "isosTriangle" => vec![("adj", 0.5)],
        "parallelogram" => vec![("adj", 0.25)],
        "trapezoid" => vec![("adj", 0.2)],
        "hexagon" => vec![("adj", 0.25)],
        "octagon" => vec![("adj", 0.292_89)],
        "rightArrow" | "leftArrow" | "upArrow" | "downArrow" => {
            vec![("adj1", 0.5), ("adj2", 0.5)]
        }
        "chevron" | "homePlate" => vec![("adj", 0.5)],
        value
            if value
                .strip_prefix("star")
                .and_then(|points| points.parse::<usize>().ok())
                .is_some_and(|points| {
                    matches!(points, 4 | 5 | 6 | 7 | 8 | 10 | 12 | 16 | 24 | 32)
                }) =>
        {
            vec![("adj", 0.45)]
        }
        _ => Vec::new(),
    };
    values
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value))
        .collect()
}

/// Adjustments use ECMA-376 guide values divided by 100000.
pub fn preset_geometry_to_path(
    shape_type: &str,
    adjustments: &HashMap<String, f64>,
    aspect_ratio: f64,
) -> Option<Vec<GeometryPathCommand>> {
    use GeometryPathCommand as C;
    let result = match shape_type {
        "rect" => vec![
            C::Move { x: 0.0, y: 0.0 },
            C::Line { x: 1.0, y: 0.0 },
            C::Line { x: 1.0, y: 1.0 },
            C::Line { x: 0.0, y: 1.0 },
            C::Close,
        ],
        "roundRect" => {
            let adjustment =
                clamp_fraction(adjustments.get("adj").copied(), ROUND_RECT_ADJUSTMENT).min(0.5);
            rounded_rect(aspect_ratio, adjustment)
        }
        "ellipse" => vec![
            C::Move { x: 1.0, y: 0.5 },
            C::Cubic {
                cp1x: 1.0,
                cp1y: 0.5 + ELLIPSE_KAPPA / 2.0,
                cp2x: 0.5 + ELLIPSE_KAPPA / 2.0,
                cp2y: 1.0,
                x: 0.5,
                y: 1.0,
            },
            C::Cubic {
                cp1x: 0.5 - ELLIPSE_KAPPA / 2.0,
                cp1y: 1.0,
                cp2x: 0.0,
                cp2y: 0.5 + ELLIPSE_KAPPA / 2.0,
                x: 0.0,
                y: 0.5,
            },
            C::Cubic {
                cp1x: 0.0,
                cp1y: 0.5 - ELLIPSE_KAPPA / 2.0,
                cp2x: 0.5 - ELLIPSE_KAPPA / 2.0,
                cp2y: 0.0,
                x: 0.5,
                y: 0.0,
            },
            C::Cubic {
                cp1x: 0.5 + ELLIPSE_KAPPA / 2.0,
                cp1y: 0.0,
                cp2x: 1.0,
                cp2y: 0.5 - ELLIPSE_KAPPA / 2.0,
                x: 1.0,
                y: 0.5,
            },
            C::Close,
        ],
        "line" | "straightConnector1" => {
            vec![C::Move { x: 0.0, y: 0.0 }, C::Line { x: 1.0, y: 1.0 }]
        }
        "triangle" | "isosTriangle" => {
            let adjustment = clamp_fraction(adjustments.get("adj").copied(), 0.5);
            polygon(&[(adjustment, 0.0), (1.0, 1.0), (0.0, 1.0)])
        }
        "rtTriangle" => polygon(&[(0.0, 0.0), (1.0, 1.0), (0.0, 1.0)]),
        "diamond" | "flowChartDecision" => {
            polygon(&[(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)])
        }
        "parallelogram" => {
            let i = clamp_fraction(adjustments.get("adj").copied(), 0.25);
            polygon(&[(i, 0.0), (1.0, 0.0), (1.0 - i, 1.0), (0.0, 1.0)])
        }
        "trapezoid" => {
            let i = clamp_fraction(adjustments.get("adj").copied(), 0.2);
            polygon(&[(i, 0.0), (1.0 - i, 0.0), (1.0, 1.0), (0.0, 1.0)])
        }
        "pentagon" | "flowChartOffpageConnector" => regular_polygon(5),
        "hexagon" => {
            let adjustment = clamp_fraction(adjustments.get("adj").copied(), 0.25).min(0.5);
            polygon(&[
                (adjustment, 0.0),
                (1.0 - adjustment, 0.0),
                (1.0, 0.5),
                (1.0 - adjustment, 1.0),
                (adjustment, 1.0),
                (0.0, 0.5),
            ])
        }
        "heptagon" => regular_polygon(7),
        "octagon" => {
            let adjustment = clamp_fraction(adjustments.get("adj").copied(), 0.292_89).min(0.5);
            polygon(&[
                (adjustment, 0.0),
                (1.0 - adjustment, 0.0),
                (1.0, adjustment),
                (1.0, 1.0 - adjustment),
                (1.0 - adjustment, 1.0),
                (adjustment, 1.0),
                (0.0, 1.0 - adjustment),
                (0.0, adjustment),
            ])
        }
        "decagon" => regular_polygon(10),
        "dodecagon" => regular_polygon(12),
        value if value.starts_with("star") => {
            let points = value[4..].parse::<usize>().ok()?;
            if !matches!(points, 4 | 5 | 6 | 7 | 8 | 10 | 12 | 16 | 24 | 32) {
                return None;
            }
            star(points, adjustments.get("adj").copied())
        }
        "bentConnector2" => bent_connector(2, adjustments.get("adj1").copied()),
        "bentConnector3" => bent_connector(3, adjustments.get("adj1").copied()),
        "bentConnector4" => bent_connector(4, adjustments.get("adj1").copied()),
        "bentConnector5" => bent_connector(5, adjustments.get("adj1").copied()),
        "curvedConnector2" => curved_connector(2),
        "curvedConnector3" => curved_connector(3),
        "curvedConnector4" => curved_connector(4),
        "curvedConnector5" => curved_connector(5),
        "rightArrow" => arrow(
            "right",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
        ),
        "leftArrow" => arrow(
            "left",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
        ),
        "upArrow" => arrow(
            "up",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
        ),
        "downArrow" => arrow(
            "down",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
        ),
        "leftRightArrow" => polygon(&[
            (0.0, 0.5),
            (0.25, 0.0),
            (0.25, 0.25),
            (0.75, 0.25),
            (0.75, 0.0),
            (1.0, 0.5),
            (0.75, 1.0),
            (0.75, 0.75),
            (0.25, 0.75),
            (0.25, 1.0),
        ]),
        "upDownArrow" => polygon(&[
            (0.5, 0.0),
            (1.0, 0.25),
            (0.75, 0.25),
            (0.75, 0.75),
            (1.0, 0.75),
            (0.5, 1.0),
            (0.0, 0.75),
            (0.25, 0.75),
            (0.25, 0.25),
            (0.0, 0.25),
        ]),
        "chevron" => {
            let notch =
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.5, aspect_ratio);
            polygon(&[
                (0.0, 0.0),
                (1.0 - notch, 0.0),
                (1.0, 0.5),
                (1.0 - notch, 1.0),
                (0.0, 1.0),
                (notch, 0.5),
            ])
        }
        "homePlate" => {
            let point =
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.5, aspect_ratio);
            polygon(&[
                (0.0, 0.0),
                (1.0 - point, 0.0),
                (1.0, 0.5),
                (1.0 - point, 1.0),
                (0.0, 1.0),
            ])
        }
        "flowChartProcess"
        | "flowChartAlternateProcess"
        | "flowChartPredefinedProcess"
        | "flowChartInternalStorage"
        | "flowChartPreparation"
        | "flowChartManualOperation"
        | "flowChartMagneticTape"
        | "flowChartMagneticDisk"
        | "flowChartMagneticDrum"
        | "flowChartDisplay"
        | "textBox" => preset_geometry_to_path("rect", adjustments, aspect_ratio)?,
        "flowChartConnector" => preset_geometry_to_path("ellipse", adjustments, aspect_ratio)?,
        "flowChartInputOutput" | "flowChartManualInput" => {
            preset_geometry_to_path("parallelogram", adjustments, aspect_ratio)?
        }
        "flowChartTerminator" => rounded_rect(aspect_ratio, 0.5),
        _ => return None,
    };
    Some(result)
}

/// Pins to `w / ss`, then converts from shortest-side to width units.
fn shortest_side_adjustment(adjustment: Option<f64>, fallback: f64, aspect_ratio: f64) -> f64 {
    let width = width_in_shortest_sides(aspect_ratio);
    pin(adjustment, fallback, width) / width
}

fn width_in_shortest_sides(aspect_ratio: f64) -> f64 {
    if aspect_ratio.is_finite() && aspect_ratio > 0.0 {
        aspect_ratio.max(1.0)
    } else {
        1.0
    }
}

fn rounded_rect(aspect_ratio: f64, adjustment: f64) -> Vec<GeometryPathCommand> {
    use GeometryPathCommand as C;
    let aspect_ratio = if aspect_ratio.is_finite() && aspect_ratio > 0.0 {
        aspect_ratio
    } else {
        1.0
    };
    let (rx, ry) = if aspect_ratio >= 1.0 {
        (adjustment / aspect_ratio, adjustment)
    } else {
        (adjustment, adjustment * aspect_ratio)
    };
    vec![
        C::Move { x: rx, y: 0.0 },
        C::Line {
            x: 1.0 - rx,
            y: 0.0,
        },
        C::Quad {
            cpx: 1.0,
            cpy: 0.0,
            x: 1.0,
            y: ry,
        },
        C::Line {
            x: 1.0,
            y: 1.0 - ry,
        },
        C::Quad {
            cpx: 1.0,
            cpy: 1.0,
            x: 1.0 - rx,
            y: 1.0,
        },
        C::Line { x: rx, y: 1.0 },
        C::Quad {
            cpx: 0.0,
            cpy: 1.0,
            x: 0.0,
            y: 1.0 - ry,
        },
        C::Line { x: 0.0, y: ry },
        C::Quad {
            cpx: 0.0,
            cpy: 0.0,
            x: rx,
            y: 0.0,
        },
        C::Close,
    ]
}

fn polygon(points: &[(f64, f64)]) -> Vec<GeometryPathCommand> {
    let mut commands = points
        .iter()
        .enumerate()
        .map(|(i, &(x, y))| {
            if i == 0 {
                GeometryPathCommand::Move { x, y }
            } else {
                GeometryPathCommand::Line { x, y }
            }
        })
        .collect::<Vec<_>>();
    if !points.is_empty() {
        commands.push(GeometryPathCommand::Close);
    }
    commands
}

fn regular_polygon(sides: usize) -> Vec<GeometryPathCommand> {
    polygon(
        &(0..sides)
            .map(|i| {
                let a = -std::f64::consts::PI / 2.0
                    + i as f64 * std::f64::consts::PI * 2.0 / sides as f64;
                (0.5 + a.cos() * 0.5, 0.5 + a.sin() * 0.5)
            })
            .collect::<Vec<_>>(),
    )
}

fn star(points: usize, adjustment: Option<f64>) -> Vec<GeometryPathCommand> {
    let inner_radius = clamp_fraction(adjustment, 0.45) * 0.5;
    polygon(
        &(0..points * 2)
            .map(|i| {
                let a =
                    -std::f64::consts::PI / 2.0 + i as f64 * std::f64::consts::PI / points as f64;
                let r = if i % 2 == 0 { 0.5 } else { inner_radius };
                (0.5 + a.cos() * r, 0.5 + a.sin() * r)
            })
            .collect::<Vec<_>>(),
    )
}

fn clamp_fraction(value: Option<f64>, fallback: f64) -> f64 {
    pin(value, fallback, 1.0)
}

fn pin(value: Option<f64>, fallback: f64, max: f64) -> f64 {
    value
        .filter(|value| value.is_finite())
        .unwrap_or(fallback)
        .clamp(0.0, max)
}

fn arrow(
    direction: &str,
    shaft_adjustment: Option<f64>,
    head_adjustment: Option<f64>,
) -> Vec<GeometryPathCommand> {
    let shaft = clamp_fraction(shaft_adjustment, 0.5);
    let edge = (1.0 - shaft) / 2.0;
    let head = clamp_fraction(head_adjustment, 0.5);
    polygon(&[
        (0.0, edge),
        (1.0 - head, edge),
        (1.0 - head, 0.0),
        (1.0, 0.5),
        (1.0 - head, 1.0),
        (1.0 - head, 1.0 - edge),
        (0.0, 1.0 - edge),
    ])
    .into_iter()
    .map(|command| match command {
        GeometryPathCommand::Move { x, y } => {
            let (x, y) = orient(direction, x, y);
            GeometryPathCommand::Move { x, y }
        }
        GeometryPathCommand::Line { x, y } => {
            let (x, y) = orient(direction, x, y);
            GeometryPathCommand::Line { x, y }
        }
        command => command,
    })
    .collect()
}

fn orient(direction: &str, x: f64, y: f64) -> (f64, f64) {
    match direction {
        "left" => (1.0 - x, y),
        "up" => (y, 1.0 - x),
        "down" => (y, x),
        _ => (x, y),
    }
}

fn bent_connector(segments: usize, adjustment: Option<f64>) -> Vec<GeometryPathCommand> {
    let bend = clamp_fraction(adjustment, 0.5);
    if segments <= 2 {
        return vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: bend, y: 0.0 },
            GeometryPathCommand::Line { x: bend, y: 1.0 },
            GeometryPathCommand::Line { x: 1.0, y: 1.0 },
        ];
    }
    let mut commands = vec![GeometryPathCommand::Move { x: 0.0, y: 0.0 }];
    for i in 1..segments {
        let fraction = i as f64 / segments as f64;
        let (x, y) = if i % 2 == 1 {
            (
                if i == 1 { bend } else { fraction },
                (i - 1) as f64 / segments as f64,
            )
        } else {
            ((i - 1) as f64 / segments as f64, fraction)
        };
        commands.push(GeometryPathCommand::Line { x, y });
    }
    commands.push(GeometryPathCommand::Line { x: 1.0, y: 1.0 });
    commands
}

fn curved_connector(segments: usize) -> Vec<GeometryPathCommand> {
    if segments <= 2 {
        return vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Cubic {
                cp1x: 0.5,
                cp1y: 0.0,
                cp2x: 0.5,
                cp2y: 1.0,
                x: 1.0,
                y: 1.0,
            },
        ];
    }
    let mut commands = vec![GeometryPathCommand::Move { x: 0.0, y: 0.0 }];
    for i in 0..segments - 1 {
        let start = i as f64 / (segments - 1) as f64;
        let end = (i + 1) as f64 / (segments - 1) as f64;
        commands.push(GeometryPathCommand::Cubic {
            cp1x: start + (end - start) * 0.5,
            cp1y: start,
            cp2x: start + (end - start) * 0.5,
            cp2y: end,
            x: end,
            y: end,
        });
    }
    commands
}

#[cfg(test)]
mod tests {
    use super::*;

    fn corner_radii(path: &[GeometryPathCommand]) -> (f64, f64) {
        let GeometryPathCommand::Move { x: rx, .. } = path[0] else {
            panic!("round rectangle must begin with a move");
        };
        let GeometryPathCommand::Quad { y: ry, .. } = path[2] else {
            panic!("round rectangle must curve its first corner");
        };
        (rx, ry)
    }

    /// Where the leading edge stops before the point begins.
    fn leading_edge(shape: &str, adjust: Option<f64>, aspect_ratio: f64) -> f64 {
        let mut adjustments = HashMap::new();
        if let Some(value) = adjust {
            adjustments.insert("adj".to_owned(), value);
        }
        let path = preset_geometry_to_path(shape, &adjustments, aspect_ratio).unwrap();
        let GeometryPathCommand::Line { x, .. } = path[1] else {
            panic!("expected a line after the opening move");
        };
        x
    }

    #[test]
    fn an_adjust_value_is_a_fraction_of_the_shortest_side() {
        assert_close(1.0 - leading_edge("chevron", Some(0.5), 1.0), 0.5);

        let aspect = 171.3 / 55.6;
        assert_close(
            1.0 - leading_edge("chevron", Some(0.5), aspect),
            0.5 / aspect,
        );

        assert_close(1.0 - leading_edge("homePlate", Some(0.5), 1.0), 0.5);
        let aspect = 280.9 / 37.5;
        assert_close(
            1.0 - leading_edge("homePlate", Some(0.5), aspect),
            0.5 / aspect,
        );
        assert_close(
            1.0 - leading_edge("homePlate", Some(0.25), aspect),
            0.25 / aspect,
        );
    }

    #[test]
    fn chevron_and_home_plate_default_to_half_the_shortest_side() {
        for shape in ["chevron", "homePlate"] {
            assert_eq!(
                preset_geometry_default_adjustments(shape)
                    .get("adj")
                    .copied(),
                Some(0.5),
                "{shape} must default to the value the spec gives it"
            );
            let aspect = 4.0;
            assert_close(1.0 - leading_edge(shape, None, aspect), 0.5 / aspect);
        }
    }

    #[test]
    fn an_adjust_value_above_half_is_honoured() {
        for shape in ["chevron", "homePlate"] {
            assert_close(1.0 - leading_edge(shape, Some(0.75), 1.0), 0.75);
            assert_close(1.0 - leading_edge(shape, Some(1.0), 1.0), 1.0);
        }
    }

    #[test]
    fn wide_shape_adjustments_may_exceed_the_shortest_side() {
        for shape in ["chevron", "homePlate"] {
            assert_close(leading_edge(shape, Some(2.0), 4.0), 0.5);
        }
        let adjustments = HashMap::from([("adj".to_owned(), 2.0)]);
        let chevron = preset_geometry_to_path("chevron", &adjustments, 4.0).unwrap();
        let GeometryPathCommand::Line { x, y } = chevron[5] else {
            panic!("expected the notch vertex");
        };
        assert_close(x, 0.5);
        assert_close(y, 0.5);
    }

    #[test]
    fn adjustments_pin_at_the_width() {
        for shape in ["chevron", "homePlate"] {
            assert_close(leading_edge(shape, Some(6.0), 4.0), 0.0);
            assert_close(leading_edge(shape, Some(2.0), 0.25), 0.0);
            assert_close(leading_edge(shape, Some(-0.25), 4.0), 1.0);
        }
    }

    #[test]
    fn normalized_adjustments_do_not_guess_raw_guide_units() {
        for shape in [
            "roundRect",
            "triangle",
            "parallelogram",
            "trapezoid",
            "hexagon",
            "octagon",
            "rightArrow",
            "star5",
            "bentConnector3",
        ] {
            let path = |value| {
                let adjustments = ["adj", "adj1", "adj2"]
                    .map(|name| (name.to_owned(), value))
                    .into();
                preset_geometry_to_path(shape, &adjustments, 1.0).unwrap()
            };
            assert_eq!(path(2.0), path(1.0), "{shape}");
        }
    }

    fn assert_close(actual: f64, expected: f64) {
        assert!((actual - expected).abs() < 1e-9, "{actual} != {expected}");
    }

    fn assert_path_close(
        actual: &[GeometryPathCommand],
        expected: &[GeometryPathCommand],
        tolerance: f64,
    ) {
        assert_eq!(actual.len(), expected.len());
        for (actual, expected) in actual.iter().zip(expected) {
            match (actual, expected) {
                (
                    GeometryPathCommand::Move { x, y },
                    GeometryPathCommand::Move {
                        x: expected_x,
                        y: expected_y,
                    },
                )
                | (
                    GeometryPathCommand::Line { x, y },
                    GeometryPathCommand::Line {
                        x: expected_x,
                        y: expected_y,
                    },
                ) => {
                    assert!((x - expected_x).abs() < tolerance);
                    assert!((y - expected_y).abs() < tolerance);
                }
                (
                    GeometryPathCommand::Quad { cpx, cpy, x, y },
                    GeometryPathCommand::Quad {
                        cpx: expected_cpx,
                        cpy: expected_cpy,
                        x: expected_x,
                        y: expected_y,
                    },
                ) => {
                    assert!((cpx - expected_cpx).abs() < tolerance);
                    assert!((cpy - expected_cpy).abs() < tolerance);
                    assert!((x - expected_x).abs() < tolerance);
                    assert!((y - expected_y).abs() < tolerance);
                }
                (GeometryPathCommand::Close, GeometryPathCommand::Close) => {}
                _ => panic!("path command variants differ"),
            }
        }
    }

    #[test]
    fn compiles_common_presets_and_rejects_unknown_shapes() {
        let adjustments = HashMap::new();
        assert!(preset_geometry_to_path("rect", &adjustments, 1.0).is_some());
        assert!(preset_geometry_to_path("ellipse", &adjustments, 1.0).is_some());
        assert!(preset_geometry_to_path("rightArrow", &adjustments, 1.0).is_some());
        assert!(preset_geometry_to_path("unknown", &adjustments, 1.0).is_none());
    }

    #[test]
    fn exposes_defaults_for_adjustable_presets() {
        assert_eq!(
            preset_geometry_default_adjustments("parallelogram").get("adj"),
            Some(&0.25)
        );
        assert_eq!(
            preset_geometry_default_adjustments("rightArrow").get("adj1"),
            Some(&0.5)
        );
        assert!(preset_geometry_default_adjustments("rect").is_empty());
    }

    #[test]
    fn non_square_round_rect_has_equal_absolute_corner_radii() {
        let path = preset_geometry_to_path("roundRect", &HashMap::new(), 4.0).unwrap();
        let (rx, ry) = corner_radii(&path);
        assert_close(rx * 400.0, ry * 100.0);
    }

    #[test]
    fn round_rect_honors_adjustment_override() {
        let path =
            preset_geometry_to_path("roundRect", &HashMap::from([("adj".to_owned(), 0.2)]), 4.0)
                .unwrap();
        let (rx, ry) = corner_radii(&path);
        assert_close(rx, 0.05);
        assert_close(ry, 0.2);
    }

    #[test]
    fn round_rect_clamps_adjustment() {
        let sharp =
            preset_geometry_to_path("roundRect", &HashMap::from([("adj".to_owned(), -0.1)]), 1.0)
                .unwrap();
        assert_eq!(corner_radii(&sharp), (0.0, 0.0));

        let pill =
            preset_geometry_to_path("roundRect", &HashMap::from([("adj".to_owned(), 0.75)]), 1.0)
                .unwrap();
        assert_eq!(corner_radii(&pill), (0.5, 0.5));
    }

    #[test]
    fn square_round_rect_matches_previous_output() {
        let path = preset_geometry_to_path("roundRect", &HashMap::new(), 1.0).unwrap();
        let previous = rounded_rect(1.0, 1.0 / 6.0);
        assert_path_close(&path, &previous, 0.000_01);
    }

    #[test]
    fn flow_chart_terminator_has_circular_ends() {
        let path = preset_geometry_to_path("flowChartTerminator", &HashMap::new(), 4.0).unwrap();
        let (rx, ry) = corner_radii(&path);
        assert_close(rx * 400.0, 50.0);
        assert_close(ry * 100.0, 50.0);
    }
}
