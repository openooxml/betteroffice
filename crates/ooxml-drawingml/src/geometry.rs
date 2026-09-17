use std::collections::HashMap;

use crate::GeometryPathCommand;

const ELLIPSE_KAPPA: f64 = 0.552_284_749_830_793_6;
const ROUND_RECT_ADJUSTMENT: f64 = 0.166_67;
/// The `vf` that puts a hexagon's corners on its frame; larger values would leave it.
const HEXAGON_VERTICAL_FACTOR: f64 = 1.154_7;

pub fn preset_geometry_default_adjustments(shape_type: &str) -> HashMap<String, f64> {
    let values = match shape_type {
        "roundRect" => vec![("adj", ROUND_RECT_ADJUSTMENT)],
        "plus" => vec![("adj", 0.25)],
        "triangle" | "isosTriangle" => vec![("adj", 0.5)],
        "parallelogram" => vec![("adj", 0.25)],
        "trapezoid" => vec![("adj", 0.25)],
        "hexagon" => vec![("adj", 0.25)],
        "octagon" => vec![("adj", 0.292_89)],
        "rightArrow" | "leftArrow" | "upArrow" | "downArrow" => {
            vec![("adj1", 0.5), ("adj2", 0.5)]
        }
        "chevron" | "homePlate" => vec![("adj", 0.5)],
        _ => star_preset(shape_type)
            .map(|star| vec![("adj", star.adjustment)])
            .unwrap_or_default(),
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
        "parallelogram" => parallelogram(shortest_side_adjustment(
            adjustments.get("adj").copied(),
            0.25,
            1.0,
            aspect_ratio,
        )),
        "plus" => plus(aspect_ratio, adjustments.get("adj").copied()),
        "trapezoid" => {
            let i =
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.25, 0.5, aspect_ratio);
            polygon(&[(i, 0.0), (1.0 - i, 0.0), (1.0, 1.0), (0.0, 1.0)])
        }
        "pentagon" | "flowChartOffpageConnector" => regular_polygon(5),
        "hexagon" => {
            let adjustment =
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.25, 0.5, aspect_ratio);
            let vertical_factor = pin(
                adjustments.get("vf").copied(),
                HEXAGON_VERTICAL_FACTOR,
                HEXAGON_VERTICAL_FACTOR,
            );
            let rise = 0.5 * vertical_factor * std::f64::consts::FRAC_PI_3.sin();
            polygon(&[
                (adjustment, 0.5 - rise),
                (1.0 - adjustment, 0.5 - rise),
                (1.0, 0.5),
                (1.0 - adjustment, 0.5 + rise),
                (adjustment, 0.5 + rise),
                (0.0, 0.5),
            ])
        }
        "heptagon" => regular_polygon(7),
        "octagon" => {
            let adjustment = pin(adjustments.get("adj").copied(), 0.292_89, 0.5);
            let x = adjustment / width_in_shortest_sides(aspect_ratio);
            let y = adjustment / height_in_shortest_sides(aspect_ratio);
            polygon(&[
                (x, 0.0),
                (1.0 - x, 0.0),
                (1.0, y),
                (1.0, 1.0 - y),
                (1.0 - x, 1.0),
                (x, 1.0),
                (0.0, 1.0 - y),
                (0.0, y),
            ])
        }
        "decagon" => regular_polygon(10),
        "dodecagon" => regular_polygon(12),
        value if value.starts_with("star") => {
            star(star_preset(value)?, adjustments.get("adj").copied())
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
            aspect_ratio,
        ),
        "leftArrow" => arrow(
            "left",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
            aspect_ratio,
        ),
        "upArrow" => arrow(
            "up",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
            aspect_ratio,
        ),
        "downArrow" => arrow(
            "down",
            adjustments.get("adj1").copied(),
            adjustments.get("adj2").copied(),
            aspect_ratio,
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
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.5, 1.0, aspect_ratio);
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
                shortest_side_adjustment(adjustments.get("adj").copied(), 0.5, 1.0, aspect_ratio);
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
        "flowChartInputOutput" | "flowChartManualInput" => parallelogram(0.25),
        "flowChartTerminator" => rounded_rect(aspect_ratio, 0.5),
        _ => return None,
    };
    Some(result)
}

/// Pins to `max · w / ss`, then converts from shortest-side to width units.
fn shortest_side_adjustment(
    adjustment: Option<f64>,
    fallback: f64,
    max: f64,
    aspect_ratio: f64,
) -> f64 {
    let width = width_in_shortest_sides(aspect_ratio);
    pin(adjustment, fallback, max * width) / width
}

fn parallelogram(offset: f64) -> Vec<GeometryPathCommand> {
    polygon(&[(offset, 0.0), (1.0, 0.0), (1.0 - offset, 1.0), (0.0, 1.0)])
}

fn height_in_shortest_sides(aspect_ratio: f64) -> f64 {
    width_in_shortest_sides(1.0 / aspect_ratio)
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

fn plus(aspect_ratio: f64, adjustment: Option<f64>) -> Vec<GeometryPathCommand> {
    let arm = pin(adjustment, 0.25, 0.5);
    let xn = arm / width_in_shortest_sides(aspect_ratio);
    let yn = arm / height_in_shortest_sides(aspect_ratio);
    polygon(&[
        (0.0, yn),
        (xn, yn),
        (xn, 0.0),
        (1.0 - xn, 0.0),
        (1.0 - xn, yn),
        (1.0, yn),
        (1.0, 1.0 - yn),
        (1.0 - xn, 1.0 - yn),
        (1.0 - xn, 1.0),
        (xn, 1.0),
        (xn, 1.0 - yn),
        (0.0, 1.0 - yn),
    ])
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

/// A `starN` preset's point count, default `adj`, and `hf`/`vf` radius factors.
#[derive(Clone, Copy)]
struct StarPreset {
    points: usize,
    adjustment: f64,
    hf: f64,
    vf: f64,
}

fn star_preset(shape_type: &str) -> Option<StarPreset> {
    let points = shape_type.strip_prefix("star")?.parse::<usize>().ok()?;
    let (adjustment, hf, vf) = match points {
        4 => (0.125, 1.0, 1.0),
        5 => (0.190_98, 1.051_46, 1.105_57),
        6 => (0.288_68, 1.154_7, 1.0),
        7 => (0.346_01, 1.025_72, 1.052_1),
        10 => (0.425_33, 1.051_46, 1.0),
        8 | 12 | 16 | 24 | 32 => (0.375, 1.0, 1.0),
        _ => return None,
    };
    Some(StarPreset {
        points,
        adjustment,
        hf,
        vf,
    })
}

fn star(preset: StarPreset, adjustment: Option<f64>) -> Vec<GeometryPathCommand> {
    let (rx, ry) = (0.5 * preset.hf, 0.5 * preset.vf);
    let inner = pin(adjustment, preset.adjustment, 0.5) * 2.0;
    polygon(
        &(0..preset.points * 2)
            .map(|i| {
                let a = -std::f64::consts::PI / 2.0
                    + i as f64 * std::f64::consts::PI / preset.points as f64;
                let scale = if i % 2 == 0 { 1.0 } else { inner };
                (0.5 + a.cos() * rx * scale, ry + a.sin() * ry * scale)
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
    aspect_ratio: f64,
) -> Vec<GeometryPathCommand> {
    let along = match direction {
        "up" | "down" => height_in_shortest_sides(aspect_ratio),
        _ => width_in_shortest_sides(aspect_ratio),
    };
    let shaft = clamp_fraction(shaft_adjustment, 0.5);
    let edge = (1.0 - shaft) / 2.0;
    let head = pin(head_adjustment, 0.5, along) / along;
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

    fn up_arrow(adj1: f64, adj2: f64, aspect_ratio: f64) -> (f64, f64) {
        let adjustments = HashMap::from([("adj1".to_owned(), adj1), ("adj2".to_owned(), adj2)]);
        let path = preset_geometry_to_path("upArrow", &adjustments, aspect_ratio).unwrap();
        let GeometryPathCommand::Line { x, y } = path[1] else {
            panic!("expected the shaft edge after the opening move");
        };
        (x, y)
    }

    #[test]
    fn an_arrow_head_uses_the_shortest_side() {
        let (edge, head) = up_arrow(0.557_13, 0.804_07, 468_000.0 / 1_078_605.0);
        assert_close(1.0 - 2.0 * edge, 0.557_13);
        assert_close(head, 0.804_07 / (1_078_605.0 / 468_000.0));
    }

    #[test]
    fn a_square_arrow_reads_its_adjustments_unchanged() {
        let (edge, head) = up_arrow(0.4, 0.7, 1.0);
        assert_close(1.0 - 2.0 * edge, 0.4);
        assert_close(head, 0.7);
        let defaults = preset_geometry_to_path("upArrow", &HashMap::new(), 1.0).unwrap();
        assert_eq!(defaults[0], GeometryPathCommand::Move { x: 0.25, y: 1.0 });
        assert_eq!(defaults[1], GeometryPathCommand::Line { x: 0.25, y: 0.5 });
    }

    #[test]
    fn an_arrow_head_pins_at_the_side_it_spans() {
        for (adjustment, expected) in [(-0.5, 0.0), (1.5, 0.375), (9.0, 1.0)] {
            let (_, head) = up_arrow(0.5, adjustment, 0.25);
            assert_close(head, expected);
        }
    }

    #[test]
    fn arrow_shafts_use_the_full_cross_axis() {
        let adjustments = HashMap::from([("adj1".to_owned(), 0.4)]);
        for (shape, aspect, x, y) in [
            ("upArrow", 4.0, 0.3, 1.0),
            ("downArrow", 4.0, 0.3, 0.0),
            ("leftArrow", 0.25, 1.0, 0.3),
            ("rightArrow", 0.25, 0.0, 0.3),
        ] {
            let path = preset_geometry_to_path(shape, &adjustments, aspect).unwrap();
            assert_eq!(path[0], GeometryPathCommand::Move { x, y }, "{shape}");
        }
    }

    #[test]
    fn arrow_heads_follow_the_pointing_axis() {
        for (shape, aspect, shoulder, tip) in [
            ("upArrow", 0.25, (0.25, 0.125), (0.5, 0.0)),
            ("downArrow", 0.25, (0.25, 0.875), (0.5, 1.0)),
            ("leftArrow", 4.0, (0.125, 0.25), (0.0, 0.5)),
            ("rightArrow", 4.0, (0.875, 0.25), (1.0, 0.5)),
        ] {
            let path = preset_geometry_to_path(shape, &HashMap::new(), aspect).unwrap();
            assert_eq!(
                path[1],
                GeometryPathCommand::Line {
                    x: shoulder.0,
                    y: shoulder.1,
                },
                "{shape}"
            );
            assert_eq!(path[3], GeometryPathCommand::Line { x: tip.0, y: tip.1 });
        }
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

    fn vertex(
        shape: &str,
        adjustments: &[(&str, f64)],
        aspect_ratio: f64,
        index: usize,
    ) -> (f64, f64) {
        let adjustments = adjustments
            .iter()
            .map(|&(name, value)| (name.to_owned(), value))
            .collect();
        match preset_geometry_to_path(shape, &adjustments, aspect_ratio).unwrap()[index] {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => (x, y),
            _ => panic!("expected a vertex"),
        }
    }

    #[test]
    fn hexagon_family_adjustments_are_fractions_of_the_shortest_side() {
        for shape in ["hexagon", "parallelogram", "trapezoid", "octagon"] {
            assert_close(vertex(shape, &[("adj", 0.25)], 4.0, 0).0, 0.0625);
            assert_close(vertex(shape, &[("adj", 0.25)], 0.25, 0).0, 0.25);
        }
        assert_close(vertex("octagon", &[("adj", 0.25)], 4.0, 2).1, 0.25);
        assert_close(vertex("octagon", &[("adj", 0.25)], 0.25, 2).1, 0.0625);
    }

    #[test]
    fn hexagon_family_adjustments_pin_at_their_aspect_scaled_maximum() {
        let aspect = 43.6;
        assert_close(
            vertex("hexagon", &[("adj", 1.29)], aspect, 0).0,
            1.29 / aspect,
        );
        assert_close(
            vertex("hexagon", &[("adj", 20.9)], aspect, 0).0,
            20.9 / aspect,
        );
        assert_close(vertex("hexagon", &[("adj", 30.0)], aspect, 0).0, 0.5);
        assert_close(
            vertex("trapezoid", &[("adj", 1.298_51)], 50.0, 0).0,
            1.298_51 / 50.0,
        );
        assert_close(vertex("trapezoid", &[("adj", 30.0)], 50.0, 0).0, 0.5);
        assert_close(vertex("parallelogram", &[("adj", 3.0)], 4.0, 0).0, 0.75);
        assert_close(vertex("parallelogram", &[("adj", 9.0)], 4.0, 0).0, 1.0);
        assert_close(vertex("octagon", &[("adj", 9.0)], 4.0, 0).0, 0.125);
    }

    #[test]
    fn trapezoid_defaults_to_a_quarter_of_the_shortest_side() {
        assert_eq!(
            preset_geometry_default_adjustments("trapezoid").get("adj"),
            Some(&0.25)
        );
        assert_close(vertex("trapezoid", &[], 4.0, 0).0, 0.0625);
    }

    #[test]
    fn hexagon_height_follows_its_vertical_factor() {
        assert!(vertex("hexagon", &[], 4.0, 0).1.abs() < 1e-6);
        assert_close(
            vertex("hexagon", &[("vf", 0.5)], 4.0, 0).1,
            0.5 - 0.25 * 3f64.sqrt() / 2.0,
        );
    }

    #[test]
    fn hexagon_vertical_factor_pins_inside_the_frame() {
        for vf in [1.2, 40.0, f64::INFINITY] {
            assert_eq!(
                vertex("hexagon", &[("vf", vf)], 4.0, 0),
                vertex("hexagon", &[], 4.0, 0),
                "{vf}"
            );
        }
        for vf in [0.0, -3.0] {
            assert_close(vertex("hexagon", &[("vf", vf)], 4.0, 0).1, 0.5);
            assert_close(vertex("hexagon", &[("vf", vf)], 4.0, 3).1, 0.5);
        }
    }

    #[test]
    fn flow_chart_input_output_ignores_the_parallelogram_adjust() {
        let adjustments = HashMap::from([("adj".to_owned(), 0.6)]);
        assert_eq!(
            preset_geometry_to_path("flowChartInputOutput", &adjustments, 4.0),
            preset_geometry_to_path("flowChartInputOutput", &HashMap::new(), 1.0),
        );
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
            "plus",
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

    const STARS: [&str; 10] = [
        "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24",
        "star32",
    ];

    fn star_vertices(shape: &str, adjust: Option<f64>) -> Vec<(f64, f64)> {
        let adjustments = adjust
            .map(|value| HashMap::from([("adj".to_owned(), value)]))
            .unwrap_or_default();
        preset_geometry_to_path(shape, &adjustments, 1.0)
            .unwrap()
            .into_iter()
            .filter_map(|command| match command {
                GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                    Some((x, y))
                }
                _ => None,
            })
            .collect()
    }

    fn cross(origin: (f64, f64), a: (f64, f64), b: (f64, f64)) -> f64 {
        (a.0 - origin.0) * (b.1 - origin.1) - (a.1 - origin.1) * (b.0 - origin.0)
    }

    #[test]
    fn five_point_star_at_its_default_is_a_regular_pentagram() {
        for adjust in [None, Some(0.190_98)] {
            let v = star_vertices("star5", adjust);
            assert!(cross(v[0], v[4], v[1]).abs() < 1e-5, "{adjust:?}");
            assert!(cross(v[0], v[4], v[3]).abs() < 1e-5, "{adjust:?}");
        }
    }

    #[test]
    fn star_inner_radius_is_twice_adj_times_the_outer() {
        let v = star_vertices("star8", Some(0.25));
        let radius = |(x, y): (f64, f64)| (x - 0.5).hypot(y - 0.5);
        assert_close(radius(v[0]), 0.5);
        assert_close(radius(v[1]), 0.25);
    }

    #[test]
    fn star_adjustment_pins_between_zero_and_half() {
        assert_eq!(
            star_vertices("star5", Some(0.8)),
            star_vertices("star5", Some(0.5))
        );
        let v = star_vertices("star8", Some(0.5));
        assert_close((v[1].0 - 0.5).hypot(v[1].1 - 0.5), 0.5);
        let v = star_vertices("star8", Some(-0.1));
        assert_close(v[1].0, 0.5);
        assert_close(v[1].1, 0.5);
    }

    #[test]
    fn stars_fill_their_frame() {
        for shape in STARS {
            let v = star_vertices(shape, None);
            let min_x = v.iter().map(|p| p.0).fold(f64::MAX, f64::min);
            let max_x = v.iter().map(|p| p.0).fold(f64::MIN, f64::max);
            let min_y = v.iter().map(|p| p.1).fold(f64::MAX, f64::min);
            let max_y = v.iter().map(|p| p.1).fold(f64::MIN, f64::max);
            for (actual, expected) in [(min_x, 0.0), (max_x, 1.0), (min_y, 0.0), (max_y, 1.0)] {
                assert!((actual - expected).abs() < 1e-4, "{shape}: {actual}");
            }
        }
    }

    #[test]
    fn stars_default_to_their_own_adjustment() {
        for (shape, expected) in [("star4", 0.125), ("star5", 0.190_98), ("star12", 0.375)] {
            assert_eq!(
                preset_geometry_default_adjustments(shape).get("adj"),
                Some(&expected)
            );
        }
        for shape in STARS {
            let default = preset_geometry_default_adjustments(shape)["adj"];
            assert_eq!(
                star_vertices(shape, None),
                star_vertices(shape, Some(default)),
                "{shape}"
            );
        }
        assert!(preset_geometry_to_path("star9", &HashMap::new(), 1.0).is_none());
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

    fn plus_path(adj: Option<f64>, aspect: f64) -> Vec<GeometryPathCommand> {
        let mut adjustments = HashMap::new();
        if let Some(value) = adj {
            adjustments.insert("adj".to_owned(), value);
        }
        preset_geometry_to_path("plus", &adjustments, aspect).unwrap()
    }

    fn plus_move(path: &[GeometryPathCommand]) -> (f64, f64) {
        let GeometryPathCommand::Move { x, y } = path[0] else {
            panic!("plus must open with a move");
        };
        (x, y)
    }

    #[test]
    fn plus_defaults_to_a_quarter_arm() {
        assert_eq!(
            preset_geometry_default_adjustments("plus").get("adj"),
            Some(&0.25)
        );
        let path = plus_path(None, 1.0);
        assert_eq!(path.len(), 13);
        assert_eq!(path[0], GeometryPathCommand::Move { x: 0.0, y: 0.25 });
        assert_eq!(path[1], GeometryPathCommand::Line { x: 0.25, y: 0.25 });
        assert_eq!(path[2], GeometryPathCommand::Line { x: 0.25, y: 0.0 });
        assert_eq!(path[5], GeometryPathCommand::Line { x: 1.0, y: 0.25 });
        assert_eq!(path[6], GeometryPathCommand::Line { x: 1.0, y: 0.75 });
        assert_eq!(path[12], GeometryPathCommand::Close);
    }

    #[test]
    fn plus_authored_adjust_matches_source_extent() {
        let adj = 39_887.0 / 100_000.0;
        let aspect = 557_530.0 / 538_480.0;
        let path = plus_path(Some(adj), aspect);
        let xn = adj / aspect;
        assert_close(plus_move(&path).1, adj);
        let GeometryPathCommand::Line { x, y } = path[1] else {
            panic!("plus second vertex carries the arm");
        };
        assert_close(x, xn);
        assert_close(y, adj);
        let GeometryPathCommand::Line { x, y } = path[6] else {
            panic!("plus right edge carries the arm");
        };
        assert_close(x, 1.0);
        assert_close(y, 1.0 - adj);
        let GeometryPathCommand::Line { x, y } = path[7] else {
            panic!("plus inner corner mirrors the arm");
        };
        assert_close(x, 1.0 - xn);
        assert_close(y, 1.0 - adj);
    }

    #[test]
    fn plus_pins_zero_and_half() {
        let (x, y) = plus_move(&plus_path(Some(0.0), 1.0));
        assert_close(x, 0.0);
        assert_close(y, 0.0);
        for pinned in [0.5, 1.0, 2.0] {
            let (x, y) = plus_move(&plus_path(Some(pinned), 1.0));
            assert_close(x, 0.0);
            assert_close(y, 0.5);
        }
        let (x, y) = plus_move(&plus_path(Some(-0.25), 1.0));
        assert_close(x, 0.0);
        assert_close(y, 0.0);
        assert_eq!(plus_path(Some(2.0), 1.0), plus_path(Some(0.5), 1.0));
        assert_eq!(plus_path(Some(-1.0), 1.0), plus_path(Some(0.0), 1.0));
    }

    #[test]
    fn plus_scales_each_axis_off_the_shortest_side() {
        let (_, y) = plus_move(&plus_path(None, 4.0));
        assert_close(y, 0.25);
        let GeometryPathCommand::Line { x, y } = plus_path(None, 4.0)[1] else {
            panic!("plus second vertex carries both axes");
        };
        assert_close(x, 0.25 / 4.0);
        assert_close(y, 0.25);
        let GeometryPathCommand::Line { x, y } = plus_path(None, 0.25)[1] else {
            panic!("tall plus mirrors the wide case");
        };
        assert_close(x, 0.25);
        assert_close(y, 0.25 * 0.25);
        let wide = plus_path(Some(0.4), 4.0);
        let tall = plus_path(Some(0.4), 0.25);
        let GeometryPathCommand::Line { x: wx, y: wy } = wide[1] else {
            unreachable!();
        };
        let GeometryPathCommand::Line { x: tx, y: ty } = tall[1] else {
            unreachable!();
        };
        assert_close(wx, 0.1);
        assert_close(wy, 0.4);
        assert_close(tx, 0.4);
        assert_close(ty, 0.1);
    }

    #[test]
    fn plus_stays_inside_a_closed_frame() {
        for aspect in [0.25, 1.0, 4.0, 557_530.0 / 538_480.0] {
            for adj in [
                None,
                Some(0.0),
                Some(0.25),
                Some(0.39887),
                Some(0.5),
                Some(2.0),
            ] {
                let path = plus_path(adj, aspect);
                assert_eq!(path.len(), 13);
                assert_eq!(path[12], GeometryPathCommand::Close);
                for command in &path {
                    match command {
                        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                            assert!((0.0..=1.0).contains(x), "{x} in {aspect} {adj:?}");
                            assert!((0.0..=1.0).contains(y), "{y} in {aspect} {adj:?}");
                        }
                        GeometryPathCommand::Close => {}
                        _ => panic!("plus uses straight edges only"),
                    }
                }
                assert_ne!(path[1], GeometryPathCommand::Line { x: 1.0, y: 0.0 });
            }
        }
    }
}
