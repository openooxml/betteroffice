use std::collections::HashMap;
use std::f64::consts::{FRAC_PI_2, PI, TAU};

use crate::GeometryPathCommand;

mod presets;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PresetPathFill {
    Normal,
    None,
    DarkenLess,
    LightenLess,
}

#[derive(Clone, Debug, PartialEq)]
pub struct PresetGeometryPath {
    pub commands: Vec<GeometryPathCommand>,
    pub fill: PresetPathFill,
    pub stroke: bool,
}

pub fn preset_geometry_layers(
    shape_type: &str,
    adjustments: &HashMap<String, f64>,
    aspect_ratio: f64,
) -> Option<Vec<PresetGeometryPath>> {
    presets::paths(shape_type, adjustments, aspect_ratio)
}

const ELLIPSE_KAPPA: f64 = 0.552_284_749_830_793_6;
const ROUND_RECT_ADJUSTMENT: f64 = 0.166_67;
/// The `vf` that puts a hexagon's corners on its frame; larger values would leave it.
const HEXAGON_VERTICAL_FACTOR: f64 = 1.154_7;
/// ECMA-376 angle units in a half turn.
const HALF_TURN_UNITS: f64 = 10_800_000.0;
const THREE_QUARTER_TURN: f64 = PI + FRAC_PI_2;
/// Float noise in the parameter span, so a half turn stays two cubic segments and not three.
const SEGMENT_SLACK: f64 = 1e-9;
/// Past this ratio a frame is degenerate, and the spec's unbounded arithmetic would overflow.
const MAX_FRAME_EXTENT: f64 = 1e12;
/// ECMA-376 leaves a callout's tail unpinned; this keeps it finite without cramping it.
const MAX_CALLOUT_OFFSET: f64 = 100.0;
/// The frame `cloudCallout`'s body path declares, in its own `w`/`h` units.
const CLOUD_FRAME: f64 = 43200.0;
/// `cloudCallout`'s body, as ECMA-376 states it: `wR`, `hR`, `stAng`, `swAng` per `arcTo`.
const CLOUD_ARCS: [[f64; 4]; 11] = [
    [6753.0, 9190.0, -11_429_249.0, 7_426_832.0],
    [5333.0, 7267.0, -8_646_143.0, 5_396_714.0],
    [4365.0, 5945.0, -8_748_475.0, 5_983_381.0],
    [4857.0, 6595.0, -7_859_164.0, 7_034_504.0],
    [5333.0, 7273.0, -4_722_533.0, 6_541_615.0],
    [6775.0, 9220.0, -2_776_035.0, 7_816_140.0],
    [5785.0, 7867.0, 37_501.0, 6_842_000.0],
    [6752.0, 9215.0, 1_347_096.0, 6_910_353.0],
    [7720.0, 10543.0, 3_974_558.0, 4_542_661.0],
    [4360.0, 5918.0, -16_496_525.0, 8_804_134.0],
    [4345.0, 5945.0, -14_809_710.0, 9_151_131.0],
];

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
        "donut" => vec![("adj", 0.25)],
        "noSmoking" => vec![("adj", 0.1875)],
        "foldedCorner" => vec![("adj", 0.166_67)],
        "corner" => vec![("adj1", 0.5), ("adj2", 0.5)],
        "mathMultiply" => vec![("adj1", 0.2352)],
        "ribbon" => vec![("adj1", 0.166_67), ("adj2", 0.5)],
        "ellipseRibbon" => vec![("adj1", 0.25), ("adj2", 0.5), ("adj3", 0.125)],
        "bentArrow" => vec![
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.4375),
        ],
        "cloudCallout" | "wedgeEllipseCallout" => vec![("adj1", -0.208_33), ("adj2", 0.625)],
        "wedgeRoundRectCallout" => vec![
            ("adj1", -0.208_33),
            ("adj2", 0.625),
            ("adj3", ROUND_RECT_ADJUSTMENT),
        ],
        _ => star_preset(shape_type)
            .map(|star| vec![("adj", star.adjustment)])
            .unwrap_or_else(|| presets::defaults(shape_type).to_vec()),
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
        "donut" => donut(Frame::new(aspect_ratio), adjustments),
        "noSmoking" => no_smoking(Frame::new(aspect_ratio), adjustments),
        "corner" => corner(Frame::new(aspect_ratio), adjustments),
        "foldedCorner" => folded_corner(Frame::new(aspect_ratio), adjustments),
        "mathMultiply" => math_multiply(Frame::new(aspect_ratio), adjustments),
        "bentArrow" => bent_arrow(Frame::new(aspect_ratio), adjustments),
        "ribbon" => ribbon(Frame::new(aspect_ratio), adjustments),
        "ellipseRibbon" => ellipse_ribbon(Frame::new(aspect_ratio), adjustments),
        "cloudCallout" => cloud_callout(Frame::new(aspect_ratio), adjustments),
        "wedgeEllipseCallout" => wedge_ellipse_callout(Frame::new(aspect_ratio), adjustments),
        "wedgeRoundRectCallout" => wedge_round_rect_callout(Frame::new(aspect_ratio), adjustments),
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
        _ => {
            let mut paths = preset_geometry_layers(shape_type, adjustments, aspect_ratio)?;
            if paths.len() != 1 {
                return None;
            }
            paths.pop()?.commands
        }
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
        C::Cubic {
            cp1x: 1.0 - rx + rx * ELLIPSE_KAPPA,
            cp1y: 0.0,
            cp2x: 1.0,
            cp2y: ry * (1.0 - ELLIPSE_KAPPA),
            x: 1.0,
            y: ry,
        },
        C::Line {
            x: 1.0,
            y: 1.0 - ry,
        },
        C::Cubic {
            cp1x: 1.0,
            cp1y: 1.0 - ry + ry * ELLIPSE_KAPPA,
            cp2x: 1.0 - rx + rx * ELLIPSE_KAPPA,
            cp2y: 1.0,
            x: 1.0 - rx,
            y: 1.0,
        },
        C::Line { x: rx, y: 1.0 },
        C::Cubic {
            cp1x: rx * (1.0 - ELLIPSE_KAPPA),
            cp1y: 1.0,
            cp2x: 0.0,
            cp2y: 1.0 - ry + ry * ELLIPSE_KAPPA,
            x: 0.0,
            y: 1.0 - ry,
        },
        C::Line { x: 0.0, y: ry },
        C::Cubic {
            cp1x: 0.0,
            cp1y: ry * (1.0 - ELLIPSE_KAPPA),
            cp2x: rx * (1.0 - ELLIPSE_KAPPA),
            cp2y: 0.0,
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
    pin_between(value, fallback, 0.0, max)
}

fn pin_between(value: Option<f64>, fallback: f64, min: f64, max: f64) -> f64 {
    unpinned(value, fallback).clamp(min, max.max(min))
}

/// An ECMA-376 adjust the spec leaves unbounded; only a non-finite value falls back.
fn unpinned(value: Option<f64>, fallback: f64) -> f64 {
    value.filter(|value| value.is_finite()).unwrap_or(fallback)
}

fn callout_offset(value: Option<f64>, fallback: f64) -> f64 {
    unpinned(value, fallback).clamp(-MAX_CALLOUT_OFFSET, MAX_CALLOUT_OFFSET)
}

fn ooxml_angle(units: f64) -> f64 {
    units * PI / HALF_TURN_UNITS
}

/// The distance from an ellipse's centre to its outline along a polar angle.
fn polar_radius(wr: f64, hr: f64, angle: f64) -> f64 {
    let extent = (hr * angle.cos()).hypot(wr * angle.sin());
    if extent > 0.0 { wr * hr / extent } else { 0.0 }
}

/// The ellipse parameter that lands on a polar angle, kept continuous across whole turns.
fn elliptical_parameter(angle: f64, wr: f64, hr: f64) -> f64 {
    if wr <= 0.0 || hr <= 0.0 {
        return angle;
    }
    let offset = (wr * angle.sin()).atan2(hr * angle.cos()) - angle;
    angle + offset - TAU * (offset / TAU).round()
}

/// The shape frame in shortest-side units, so ECMA-376's `ss` is 1 and `w`/`h` are at least 1.
#[derive(Clone, Copy)]
struct Frame {
    w: f64,
    h: f64,
}

impl Frame {
    fn new(aspect_ratio: f64) -> Self {
        Self {
            w: width_in_shortest_sides(aspect_ratio).min(MAX_FRAME_EXTENT),
            h: height_in_shortest_sides(aspect_ratio).min(MAX_FRAME_EXTENT),
        }
    }
}

/// Transcribes an ECMA-376 `pathLst` in frame units, normalizing each point to the unit frame.
struct Outline {
    frame: Frame,
    commands: Vec<GeometryPathCommand>,
    cursor: (f64, f64),
}

impl Outline {
    fn new(frame: Frame) -> Self {
        Self {
            frame,
            commands: Vec::new(),
            cursor: (0.0, 0.0),
        }
    }

    fn at(&self, x: f64, y: f64) -> (f64, f64) {
        (x / self.frame.w, y / self.frame.h)
    }

    fn move_to(mut self, x: f64, y: f64) -> Self {
        self.cursor = (x, y);
        let (x, y) = self.at(x, y);
        self.commands.push(GeometryPathCommand::Move { x, y });
        self
    }

    fn line_to(mut self, x: f64, y: f64) -> Self {
        self.cursor = (x, y);
        let (x, y) = self.at(x, y);
        self.commands.push(GeometryPathCommand::Line { x, y });
        self
    }

    fn quad_to(mut self, cpx: f64, cpy: f64, x: f64, y: f64) -> Self {
        self.cursor = (x, y);
        let (cpx, cpy) = self.at(cpx, cpy);
        let (x, y) = self.at(x, y);
        self.commands
            .push(GeometryPathCommand::Quad { cpx, cpy, x, y });
        self
    }

    fn cubic_to(mut self, cp1x: f64, cp1y: f64, cp2x: f64, cp2y: f64, x: f64, y: f64) -> Self {
        self.cursor = (x, y);
        let (cp1x, cp1y) = self.at(cp1x, cp1y);
        let (cp2x, cp2y) = self.at(cp2x, cp2y);
        let (x, y) = self.at(x, y);
        self.commands.push(GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        });
        self
    }

    /// ECMA-376 `arcTo`: `stAng`/`swAng` are polar angles, and the ellipse is centred so that
    /// `start` lands on the pen position.
    fn arc_to(mut self, wr: f64, hr: f64, start: f64, sweep: f64) -> Self {
        let first = elliptical_parameter(start, wr, hr);
        let last = elliptical_parameter(start + sweep, wr, hr);
        let cx = self.cursor.0 - wr * first.cos();
        let cy = self.cursor.1 - hr * first.sin();
        let forward = last >= first;
        let mut from = first;
        for step in 0..8 {
            let quadrant = from / FRAC_PI_2;
            let to = if step == 7 {
                last
            } else if forward {
                (((quadrant + SEGMENT_SLACK).floor() + 1.0) * FRAC_PI_2).min(last)
            } else {
                (((quadrant - SEGMENT_SLACK).ceil() - 1.0) * FRAC_PI_2).max(last)
            };
            let alpha = 4.0 / 3.0 * ((to - from) / 4.0).tan();
            let start_point = (cx + wr * from.cos(), cy + hr * from.sin());
            let end_point = (cx + wr * to.cos(), cy + hr * to.sin());
            self = self.cubic_to(
                start_point.0 - alpha * wr * from.sin(),
                start_point.1 + alpha * hr * from.cos(),
                end_point.0 + alpha * wr * to.sin(),
                end_point.1 - alpha * hr * to.cos(),
                end_point.0,
                end_point.1,
            );
            if to == last {
                break;
            }
            from = to;
        }
        self
    }

    fn close(mut self) -> Self {
        self.commands.push(GeometryPathCommand::Close);
        self
    }

    fn finish(self) -> Vec<GeometryPathCommand> {
        self.commands
    }
}

fn frame_polygon(frame: Frame, points: &[(f64, f64)]) -> Vec<GeometryPathCommand> {
    polygon(
        &points
            .iter()
            .map(|&(x, y)| (x / frame.w, y / frame.h))
            .collect::<Vec<_>>(),
    )
}

fn donut(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (wd2, hd2) = (frame.w / 2.0, frame.h / 2.0);
    let dr = pin(adjustments.get("adj").copied(), 0.25, 0.5);
    Outline::new(frame)
        .move_to(0.0, hd2)
        .arc_to(wd2, hd2, PI, TAU)
        .close()
        .move_to(dr, hd2)
        .arc_to(wd2 - dr, hd2 - dr, PI, -TAU)
        .close()
        .finish()
}

fn no_smoking(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (wd2, hd2) = (frame.w / 2.0, frame.h / 2.0);
    let dr = pin(adjustments.get("adj").copied(), 0.1875, 0.5);
    let (iwd2, ihd2) = (wd2 - dr, hd2 - dr);
    let diagonal = frame.h.atan2(frame.w);
    let half_gap = (dr / 2.0).atan2(polar_radius(iwd2, ihd2, diagonal));
    let (start, sweep) = (diagonal - half_gap, 2.0 * half_gap - PI);
    let reach = polar_radius(iwd2, ihd2, start);
    let (dx, dy) = (reach * start.cos(), reach * start.sin());
    Outline::new(frame)
        .move_to(0.0, hd2)
        .arc_to(wd2, hd2, PI, TAU)
        .close()
        .move_to(wd2 + dx, hd2 + dy)
        .arc_to(iwd2, ihd2, start, sweep)
        .close()
        .move_to(wd2 - dx, hd2 - dy)
        .arc_to(iwd2, ihd2, start - PI, sweep)
        .close()
        .finish()
}

fn corner(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let leg = pin(adjustments.get("adj1").copied(), 0.5, h);
    let x1 = pin(adjustments.get("adj2").copied(), 0.5, w);
    let y1 = h - leg;
    frame_polygon(
        frame,
        &[(0.0, 0.0), (x1, 0.0), (x1, y1), (w, y1), (w, h), (0.0, h)],
    )
}

fn folded_corner(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dy2 = pin(adjustments.get("adj").copied(), ROUND_RECT_ADJUSTMENT, 0.5);
    let dy1 = dy2 / 5.0;
    let (x1, y2) = (w - dy2, h - dy2);
    let (x2, y1) = (x1 + dy1, y2 + dy1);
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(w, 0.0)
        .line_to(w, y2)
        .line_to(x1, h)
        .line_to(0.0, h)
        .close()
        .move_to(x1, h)
        .line_to(x2, y1)
        .line_to(w, y2)
        .close()
        .finish()
}

fn math_multiply(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    const ARM_FRACTION: f64 = 0.519_65;
    let (w, h) = (frame.w, frame.h);
    let (hc, vc) = (w / 2.0, h / 2.0);
    let th = pin(adjustments.get("adj1").copied(), 0.2352, ARM_FRACTION);
    let angle = h.atan2(w);
    let (sa, ca, ta) = (angle.sin(), angle.cos(), angle.tan());
    let arm = w.hypot(h) * (1.0 - ARM_FRACTION);
    let (xm, ym) = (ca * arm / 2.0, sa * arm / 2.0);
    let (dx, dy) = (sa * th / 2.0, ca * th / 2.0);
    let (xa, ya) = (xm - dx, ym + dy);
    let (xb, yb) = (xm + dx, ym - dy);
    let yc = (hc - xb) * ta + yb;
    let (xd, xe) = (w - xb, w - xa);
    let reach = (vc - ya) / ta;
    let (xf, xl) = (xe - reach, xa + reach);
    let (yg, yh, yi) = (h - ya, h - yb, h - yc);
    frame_polygon(
        frame,
        &[
            (xa, ya),
            (xb, yb),
            (hc, yc),
            (xd, yb),
            (xe, ya),
            (xf, vc),
            (xe, yg),
            (xd, yh),
            (hc, yi),
            (xb, yh),
            (xa, yg),
            (xl, vc),
        ],
    )
}

fn bent_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let aw2 = pin(adjustments.get("adj2").copied(), 0.25, 0.5);
    let th = pin(adjustments.get("adj1").copied(), 0.25, 2.0 * aw2);
    let ah = pin(adjustments.get("adj3").copied(), 0.25, 0.5);
    let dh2 = aw2 - th / 2.0;
    let bend = pin(
        adjustments.get("adj4").copied(),
        0.4375,
        (w - ah).min(h - dh2),
    );
    let inner = (bend - th).max(0.0);
    let (x3, x4) = (th + inner, w - ah);
    let y3 = dh2 + th;
    let (y4, y5) = (y3 + dh2, dh2 + bend);
    Outline::new(frame)
        .move_to(0.0, h)
        .line_to(0.0, y5)
        .arc_to(bend, bend, PI, FRAC_PI_2)
        .line_to(x4, dh2)
        .line_to(x4, 0.0)
        .line_to(w, aw2)
        .line_to(x4, y4)
        .line_to(x4, y3)
        .line_to(x3, y3)
        .arc_to(inner, inner, THREE_QUARTER_TURN, -FRAC_PI_2)
        .line_to(th, h)
        .close()
        .finish()
}

fn ribbon(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = pin(
        adjustments.get("adj1").copied(),
        ROUND_RECT_ADJUSTMENT,
        0.333_33,
    );
    let a2 = pin_between(adjustments.get("adj2").copied(), 0.5, 0.25, 0.75);
    let (wd8, wd32) = (w / 8.0, w / 32.0);
    let dx2 = w * a2 / 2.0;
    let (x2, x9) = (w / 2.0 - dx2, w / 2.0 + dx2);
    let (x3, x8) = (x2 + wd32, x9 - wd32);
    let (x5, x6) = (x2 + wd8, x9 - wd8);
    let (x4, x7) = (x5 - wd32, x6 + wd32);
    let (y1, y2) = (h * a1 / 2.0, h * a1);
    let y4 = h - y2;
    let y3 = y4 / 2.0;
    let hr = h * a1 / 4.0;
    let y5 = h - hr;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(x4, 0.0)
        .arc_to(wd32, hr, THREE_QUARTER_TURN, PI)
        .line_to(x3, y1)
        .arc_to(wd32, hr, THREE_QUARTER_TURN, -PI)
        .line_to(x8, y2)
        .arc_to(wd32, hr, FRAC_PI_2, -PI)
        .line_to(x7, y1)
        .arc_to(wd32, hr, FRAC_PI_2, PI)
        .line_to(w, 0.0)
        .line_to(w - wd8, y3)
        .line_to(w, y4)
        .line_to(x9, y4)
        .line_to(x9, y5)
        .arc_to(wd32, hr, 0.0, FRAC_PI_2)
        .line_to(x3, h)
        .arc_to(wd32, hr, FRAC_PI_2, FRAC_PI_2)
        .line_to(x2, y4)
        .line_to(0.0, y4)
        .line_to(wd8, y3)
        .close()
        .finish()
}

fn ellipse_ribbon(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let (hc, wd8) = (w / 2.0, w / 8.0);
    let a1 = clamp_fraction(adjustments.get("adj1").copied(), 0.25);
    let a2 = pin_between(adjustments.get("adj2").copied(), 0.5, 0.25, 0.75);
    let a3 = pin_between(
        adjustments.get("adj3").copied(),
        0.125,
        (a1 - (1.0 - a1) / 2.0).max(0.0),
        a1,
    );
    let x2 = hc - w * a2 / 2.0;
    let x3 = x2 + wd8;
    let (x4, x5, x6) = (w - x3, w - x2, w - wd8);
    let dy1 = h * a3;
    let slope = 4.0 * dy1 / w;
    let rise = |x: f64| slope * (x - x * x / w);
    let y1 = rise(x3);
    let (cx1, cy1) = (x3 / 2.0, slope * x3 / 2.0);
    let cx2 = w - cx1;
    let depth = h * a1;
    let dy3 = depth - dy1;
    let tail_rise = rise(x2);
    let y3 = tail_rise + dy3;
    let cy3 = 2.0 * dy1 + dy3 - tail_rise;
    let rh = h - depth;
    let y2 = (dy1 * 14.0 / 16.0 + rh) / 2.0;
    let (y5, y6) = (tail_rise + rh, y3 + rh);
    let cx4 = x2 / 2.0;
    let (cy4, cx5) = (slope * cx4 + rh, w - cx4);
    let cy6 = cy3 + rh;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .quad_to(cx1, cy1, x3, y1)
        .line_to(x2, y3)
        .quad_to(hc, cy3, x5, y3)
        .line_to(x4, y1)
        .quad_to(cx2, cy1, w, 0.0)
        .line_to(x6, y2)
        .line_to(w, rh)
        .quad_to(cx5, cy4, x5, y5)
        .line_to(x5, y6)
        .quad_to(hc, cy6, x2, y6)
        .line_to(x2, y5)
        .quad_to(cx4, cy4, 0.0, rh)
        .line_to(wd8, y2)
        .close()
        .finish()
}

fn cloud_callout(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let square = Frame {
        w: CLOUD_FRAME,
        h: CLOUD_FRAME,
    };
    let mut body = Outline::new(square).move_to(3900.0, 14370.0);
    for [wr, hr, start, sweep] in CLOUD_ARCS {
        body = body.arc_to(wr, hr, ooxml_angle(start), ooxml_angle(sweep));
    }
    let mut commands = body.close().finish();
    let (hc, vc) = (w / 2.0, h / 2.0);
    let dx_pos = w * callout_offset(adjustments.get("adj1").copied(), -0.208_33);
    let dy_pos = h * callout_offset(adjustments.get("adj2").copied(), 0.625);
    let (x_pos, y_pos) = (hc + dx_pos, vc + dy_pos);
    let pointer = dy_pos.atan2(dx_pos);
    let edge = (hc * pointer.sin()).atan2(vc * pointer.cos());
    let g6 = hc + hc * edge.cos() - x_pos;
    let g7 = vc + vc * edge.sin() - y_pos;
    let span = g6.hypot(g7);
    let (ux, uy) = if span > 0.0 {
        (g6 / span, g7 / span)
    } else {
        (0.0, 0.0)
    };
    let third = (span - 6600.0 / 21600.0) / 3.0;
    let (small, medium, large) = (600.0 / 21600.0, 1200.0 / 21600.0, 1800.0 / 21600.0);
    let near = third + large;
    let far = 4800.0 / 21600.0 + 2.0 * third;
    let bubbles = [
        (x_pos, y_pos, small),
        (x_pos + near * ux, y_pos + near * uy, medium),
        (x_pos + far * ux, y_pos + far * uy, large),
    ];
    for (cx, cy, r) in bubbles {
        commands.extend(
            Outline::new(frame)
                .move_to(cx + r, cy)
                .arc_to(r, r, 0.0, TAU)
                .close()
                .finish(),
        );
    }
    commands
}

fn wedge_ellipse_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    const TAIL_HALF_WIDTH: f64 = 660_000.0;
    let (wd2, hd2) = (frame.w / 2.0, frame.h / 2.0);
    let dx_pos = frame.w * callout_offset(adjustments.get("adj1").copied(), -0.208_33);
    let dy_pos = frame.h * callout_offset(adjustments.get("adj2").copied(), 0.625);
    let pointer = (dy_pos * frame.w).atan2(dx_pos * frame.h);
    let gap = ooxml_angle(TAIL_HALF_WIDTH);
    let on_edge = |angle: f64| (wd2 * angle.cos(), hd2 * angle.sin());
    let (dx1, dy1) = on_edge(pointer + gap);
    let (dx2, dy2) = on_edge(pointer - gap);
    let (lead, trail) = (dy1.atan2(dx1), dy2.atan2(dx2));
    let sweep = if trail > lead {
        trail - lead
    } else {
        trail - lead + TAU
    };
    Outline::new(frame)
        .move_to(wd2 + dx_pos, hd2 + dy_pos)
        .line_to(wd2 + dx1, hd2 + dy1)
        .arc_to(wd2, hd2, lead, sweep)
        .close()
        .finish()
}

fn wedge_round_rect_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let (hc, vc) = (w / 2.0, h / 2.0);
    let dx_pos = w * callout_offset(adjustments.get("adj1").copied(), -0.208_33);
    let dy_pos = h * callout_offset(adjustments.get("adj2").copied(), 0.625);
    let (x_pos, y_pos) = (hc + dx_pos, vc + dy_pos);
    let steeper = dy_pos.abs() - (dx_pos * h / w).abs() > 0.0;
    let pick = |flag: bool, when: f64, otherwise: f64| if flag { when } else { otherwise };
    let x1 = w * pick(dx_pos > 0.0, 7.0, 2.0) / 12.0;
    let x2 = w * pick(dx_pos > 0.0, 10.0, 5.0) / 12.0;
    let y1 = h * pick(dy_pos > 0.0, 7.0, 2.0) / 12.0;
    let y2 = h * pick(dy_pos > 0.0, 10.0, 5.0) / 12.0;
    let xl = pick(steeper, 0.0, pick(dx_pos > 0.0, 0.0, x_pos));
    let xt = pick(steeper, pick(dy_pos > 0.0, x1, x_pos), x1);
    let xr = pick(steeper, w, pick(dx_pos > 0.0, x_pos, w));
    let xb = pick(steeper, pick(dy_pos > 0.0, x_pos, x1), x1);
    let yl = pick(steeper, y1, pick(dx_pos > 0.0, y1, y_pos));
    let yt = pick(steeper, pick(dy_pos > 0.0, 0.0, y_pos), 0.0);
    let yr = pick(steeper, y1, pick(dx_pos > 0.0, y_pos, y1));
    let yb = pick(steeper, pick(dy_pos > 0.0, y_pos, h), h);
    let u1 = pin(adjustments.get("adj3").copied(), ROUND_RECT_ADJUSTMENT, 0.5);
    Outline::new(frame)
        .move_to(0.0, u1)
        .arc_to(u1, u1, PI, FRAC_PI_2)
        .line_to(x1, 0.0)
        .line_to(xt, yt)
        .line_to(x2, 0.0)
        .line_to(w - u1, 0.0)
        .arc_to(u1, u1, THREE_QUARTER_TURN, FRAC_PI_2)
        .line_to(w, y1)
        .line_to(xr, yr)
        .line_to(w, y2)
        .line_to(w, h - u1)
        .arc_to(u1, u1, 0.0, FRAC_PI_2)
        .line_to(x2, h)
        .line_to(xb, yb)
        .line_to(x1, h)
        .line_to(u1, h)
        .arc_to(u1, u1, FRAC_PI_2, FRAC_PI_2)
        .line_to(0.0, y2)
        .line_to(xl, yl)
        .line_to(0.0, y1)
        .close()
        .finish()
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
        let GeometryPathCommand::Cubic { y: ry, .. } = path[2] else {
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
    fn round_rect_corners_follow_circular_arcs() {
        let path = preset_geometry_to_path("roundRect", &HashMap::new(), 1.0).unwrap();
        let GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } = path[2]
        else {
            panic!("expected a circular corner");
        };
        let radius = ROUND_RECT_ADJUSTMENT;
        let midpoint_x = ((1.0 - radius) + 3.0 * cp1x + 3.0 * cp2x + x) / 8.0;
        let midpoint_y = (3.0 * cp1y + 3.0 * cp2y + y) / 8.0;
        assert_close(
            (midpoint_x - (1.0 - radius)).hypot(midpoint_y - radius),
            radius,
        );
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
