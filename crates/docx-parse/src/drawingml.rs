//! Bug-compatible subset of `@openooxml/drawingml` used by DOCX drawings.
//!
//! Despite its name, the incumbent `findByFullName` helper only searches
//! direct children. This module deliberately retains that behavior, together
//! with JavaScript `parseInt` prefix parsing and explicit null arrow fields.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use crate::scalars::ColorValue;
use crate::xml::{XmlElement, parse_javascript_integer_prefix};

const ELLIPSE_KAPPA: f64 = 0.552_284_749_830_793_6;
const ROUND_RECT_RADIUS: f64 = 0.166_666_666_666_666_66;
const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const MAX_CUSTOM_PATH_COMMANDS: usize = 2_048;
const MAX_CUSTOM_GUIDES: usize = 512;
const MAX_ABS_CUSTOM_PATH_NUMBER: f64 = 1_000_000_000.0;
const MAX_ABS_NORMALIZED_PATH_NUMBER: f64 = 10_000.0;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeFill {
    #[serde(rename = "type")]
    pub fill_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFill>,
}

impl ShapeFill {
    fn named(fill_type: &str) -> Self {
        Self {
            fill_type: fill_type.to_owned(),
            color: None,
            gradient: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct GradientFill {
    #[serde(rename = "type")]
    pub gradient_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub angle: Option<f64>,
    pub stops: Vec<GradientStop>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct GradientStop {
    pub position: f64,
    pub color: ColorValue,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeOutline {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub join: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub head_end: Option<LineEnd>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<LineEnd>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct LineEnd {
    #[serde(rename = "type")]
    pub end_type: String,
    // Not skipped: getAttribute returns null and the TS object retains it.
    pub width: Option<String>,
    pub length: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum GeometryPathCommand {
    Move {
        x: f64,
        y: f64,
    },
    Line {
        x: f64,
        y: f64,
    },
    Quad {
        cpx: f64,
        cpy: f64,
        x: f64,
        y: f64,
    },
    Cubic {
        cp1x: f64,
        cp1y: f64,
        cp2x: f64,
        cp2y: f64,
        x: f64,
        y: f64,
    },
    Close,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct Size2D {
    pub width: f64,
    pub height: f64,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transform2D {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Point2D {
    pub x: f64,
    pub y: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TransformResult {
    pub size: Size2D,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<Transform2D>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<Point2D>,
}

pub fn parse_color_element(element: Option<&XmlElement>) -> Option<ColorValue> {
    let element = element?;
    if let Some(color) = element.child_by_full_name("a:srgbClr") {
        if let Some(value) = color.attribute(None, "val") {
            return Some(apply_color_modifiers(
                ColorValue {
                    rgb: Some(value.to_owned()),
                    ..ColorValue::default()
                },
                color,
            ));
        }
    }
    if let Some(color) = element.child_by_full_name("a:schemeClr") {
        if let Some(value) = color.attribute(None, "val") {
            return Some(apply_color_modifiers(
                ColorValue {
                    theme_color: Some(scheme_color(value).to_owned()),
                    ..ColorValue::default()
                },
                color,
            ));
        }
    }
    if let Some(color) = element.child_by_full_name("a:sysClr") {
        return Some(ColorValue {
            rgb: Some(
                color
                    .attribute(None, "lastClr")
                    .unwrap_or("000000")
                    .to_owned(),
            ),
            ..ColorValue::default()
        });
    }
    if let Some(color) = element.child_by_full_name("a:prstClr") {
        let rgb = match color.attribute(None, "val")? {
            "black" => "000000",
            "white" => "FFFFFF",
            "red" => "FF0000",
            "green" => "00FF00",
            "blue" => "0000FF",
            "yellow" => "FFFF00",
            "cyan" => "00FFFF",
            "magenta" => "FF00FF",
            _ => return None,
        };
        return Some(ColorValue {
            rgb: Some(rgb.to_owned()),
            ..ColorValue::default()
        });
    }
    None
}

fn apply_color_modifiers(mut color: ColorValue, element: &XmlElement) -> ColorValue {
    if let Some(value) = element
        .child_by_full_name("a:shade")
        .and_then(|child| child.attribute(None, "val"))
        .and_then(parse_javascript_integer_prefix)
    {
        color.theme_shade = Some(format!(
            "{:02X}",
            js_round(value / 100_000.0 * 255.0) as i64
        ));
    }
    if let Some(value) = element
        .child_by_full_name("a:tint")
        .and_then(|child| child.attribute(None, "val"))
        .and_then(parse_javascript_integer_prefix)
    {
        color.theme_tint = Some(format!(
            "{:02X}",
            js_round(value / 100_000.0 * 255.0) as i64
        ));
    }
    color
}

fn scheme_color(value: &str) -> &str {
    match value {
        "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "dk1" | "lt1"
        | "dk2" | "lt2" | "hlink" | "folHlink" => value,
        "tx1" => "text1",
        "tx2" => "text2",
        "bg1" => "background1",
        "bg2" => "background2",
        _ => "dk1",
    }
}

pub fn parse_fill(sp_pr: Option<&XmlElement>) -> Option<ShapeFill> {
    let sp_pr = sp_pr?;
    if sp_pr.child_by_full_name("a:noFill").is_some() {
        return Some(ShapeFill::named("none"));
    }
    if let Some(solid) = sp_pr.child_by_full_name("a:solidFill") {
        return Some(ShapeFill {
            fill_type: "solid".to_owned(),
            color: parse_color_element(Some(solid)),
            gradient: None,
        });
    }
    if sp_pr.child_by_full_name("a:gradFill").is_some() {
        return Some(ShapeFill::named("gradient"));
    }
    None
}

pub fn parse_outline(sp_pr: Option<&XmlElement>) -> Option<ShapeOutline> {
    let line = sp_pr?.child_by_full_name("a:ln")?;
    if line.child_by_full_name("a:noFill").is_some() {
        return None;
    }
    Some(ShapeOutline {
        width: numeric_attr(Some(line), "w"),
        color: line
            .child_by_full_name("a:solidFill")
            .and_then(|fill| parse_color_element(Some(fill))),
        style: line
            .child_by_full_name("a:prstDash")
            .and_then(|dash| dash.attribute(None, "val"))
            .map(str::to_owned),
        ..ShapeOutline::default()
    })
}

pub fn resolve_color_value_to_hex(color: Option<&ColorValue>) -> Option<String> {
    let color = color?;
    if let Some(rgb) = &color.rgb {
        return Some(format!("#{rgb}"));
    }
    color
        .theme_color
        .as_deref()
        .map(default_theme_color)
        .map(|rgb| format!("#{rgb}"))
}

fn default_theme_color(slot: &str) -> &str {
    match slot {
        "dk1" | "text1" => "000000",
        "lt1" | "background1" => "FFFFFF",
        "dk2" | "text2" => "44546A",
        "lt2" | "background2" => "E7E6E6",
        "accent1" => "4472C4",
        "accent2" => "ED7D31",
        "accent3" => "A5A5A5",
        "accent4" => "FFC000",
        "accent5" => "5B9BD5",
        "accent6" => "70AD47",
        "hlink" => "0563C1",
        "folHlink" => "954F72",
        _ => "000000",
    }
}

pub fn parse_gradient_fill(element: &XmlElement) -> ShapeFill {
    let mut gradient_type = "linear";
    let mut angle = None;
    if let Some(linear) = element.child_by_full_name("a:lin") {
        angle = numeric_attr(Some(linear), "ang").map(|value| value / 60_000.0);
    }
    if let Some(path) = element.child_by_full_name("a:path") {
        gradient_type = match path.attribute(None, "path") {
            Some("circle") => "radial",
            Some("rect") => "rectangular",
            _ => "path",
        };
    }
    let stops = element
        .child_by_full_name("a:gsLst")
        .into_iter()
        .flat_map(|list| list.children_by_local_name("gs"))
        .filter_map(|stop| {
            Some(GradientStop {
                position: numeric_attr(Some(stop), "pos").unwrap_or(0.0),
                color: parse_color_element(Some(stop))?,
            })
        })
        .collect();
    ShapeFill {
        fill_type: "gradient".to_owned(),
        color: None,
        gradient: Some(GradientFill {
            gradient_type: gradient_type.to_owned(),
            angle,
            stops,
        }),
    }
}

pub fn parse_line_end(element: &XmlElement) -> LineEnd {
    let end_type = match element.attribute(None, "type").unwrap_or("none") {
        value @ ("none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow") => value,
        _ => "none",
    };
    LineEnd {
        end_type: end_type.to_owned(),
        width: element.attribute(None, "w").map(str::to_owned),
        length: element.attribute(None, "len").map(str::to_owned),
    }
}

pub fn rot_to_degrees(value: Option<&str>) -> Option<f64> {
    let value = value?;
    if value.is_empty() {
        return None;
    }
    parse_javascript_integer_prefix(value).map(|value| value / 60_000.0)
}

pub fn parse_transform(element: Option<&XmlElement>) -> TransformResult {
    let Some(element) = element else {
        return TransformResult {
            size: Size2D::default(),
            transform: None,
            offset: None,
        };
    };
    let extent = element.child_by_full_name("a:ext");
    let size = Size2D {
        width: numeric_attr(extent, "cx").unwrap_or(0.0),
        height: numeric_attr(extent, "cy").unwrap_or(0.0),
    };
    let offset = element.child_by_full_name("a:off").map(|offset| Point2D {
        x: numeric_attr(Some(offset), "x").unwrap_or(0.0),
        y: numeric_attr(Some(offset), "y").unwrap_or(0.0),
    });
    let rotation = rot_to_degrees(element.attribute(None, "rot"));
    let flip_h = (element.attribute(None, "flipH") == Some("1")).then_some(true);
    let flip_v = (element.attribute(None, "flipV") == Some("1")).then_some(true);
    let transform =
        (rotation.is_some() || flip_h.is_some() || flip_v.is_some()).then_some(Transform2D {
            rotation,
            flip_h,
            flip_v,
        });
    TransformResult {
        size,
        transform,
        offset,
    }
}

pub fn parse_shape_type(sp_pr: Option<&XmlElement>) -> String {
    sp_pr
        .and_then(|properties| properties.child_by_full_name("a:prstGeom"))
        .and_then(|preset| preset.attribute(None, "prst"))
        .unwrap_or("rect")
        .to_owned()
}

pub fn parse_preset_geometry_path(sp_pr: Option<&XmlElement>) -> Option<Vec<GeometryPathCommand>> {
    let preset = sp_pr?.child_by_full_name("a:prstGeom")?;
    let shape_type = preset.attribute(None, "prst")?;
    let mut values = standard_guide_values(100_000.0, 100_000.0);
    apply_guide_list(Some(preset), "avLst", &mut values);
    let mut adjustments = HashMap::new();
    if let Some(list) = local_child(Some(preset), "avLst") {
        for guide in list.children_by_local_name("gd").take(MAX_CUSTOM_GUIDES) {
            if let Some(name) = guide.attribute(None, "name") {
                if let Some(value) = values.get(name) {
                    adjustments.insert(name.to_owned(), *value);
                }
            }
        }
    }
    preset_geometry_to_path(shape_type, &adjustments)
}

pub fn preset_geometry_to_path(
    shape_type: &str,
    adjustments: &HashMap<String, f64>,
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
            let r = ROUND_RECT_RADIUS;
            vec![
                C::Move { x: r, y: 0.0 },
                C::Line { x: 1.0 - r, y: 0.0 },
                C::Quad {
                    cpx: 1.0,
                    cpy: 0.0,
                    x: 1.0,
                    y: r,
                },
                C::Line { x: 1.0, y: 1.0 - r },
                C::Quad {
                    cpx: 1.0,
                    cpy: 1.0,
                    x: 1.0 - r,
                    y: 1.0,
                },
                C::Line { x: r, y: 1.0 },
                C::Quad {
                    cpx: 0.0,
                    cpy: 1.0,
                    x: 0.0,
                    y: 1.0 - r,
                },
                C::Line { x: 0.0, y: r },
                C::Quad {
                    cpx: 0.0,
                    cpy: 0.0,
                    x: r,
                    y: 0.0,
                },
                C::Close,
            ]
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
        "triangle" | "isosTriangle" => polygon(&[(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)]),
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
        "hexagon" => regular_polygon(6),
        "heptagon" => regular_polygon(7),
        "octagon" => regular_polygon(8),
        "decagon" => regular_polygon(10),
        "dodecagon" => regular_polygon(12),
        value if value.starts_with("star") => {
            let points = value[4..].parse::<usize>().ok()?;
            if !matches!(points, 4 | 5 | 6 | 7 | 8 | 10 | 12 | 16 | 24 | 32) {
                return None;
            }
            star(points)
        }
        "bentConnector2" => bent_connector(2, adjustments.get("adj1").copied()),
        "bentConnector3" => bent_connector(3, adjustments.get("adj1").copied()),
        "bentConnector4" => bent_connector(4, adjustments.get("adj1").copied()),
        "bentConnector5" => bent_connector(5, adjustments.get("adj1").copied()),
        "curvedConnector2" => curved_connector(2),
        "curvedConnector3" => curved_connector(3),
        "curvedConnector4" => curved_connector(4),
        "curvedConnector5" => curved_connector(5),
        "rightArrow" => arrow("right", adjustments.get("adj2").copied()),
        "leftArrow" => arrow("left", adjustments.get("adj2").copied()),
        "upArrow" => arrow("up", adjustments.get("adj2").copied()),
        "downArrow" => arrow("down", adjustments.get("adj2").copied()),
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
        "chevron" => polygon(&[
            (0.0, 0.0),
            (0.65, 0.0),
            (1.0, 0.5),
            (0.65, 1.0),
            (0.0, 1.0),
            (0.35, 0.5),
        ]),
        "homePlate" => polygon(&[(0.0, 0.0), (0.75, 0.0), (1.0, 0.5), (0.75, 1.0), (0.0, 1.0)]),
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
        | "textBox" => preset_geometry_to_path("rect", adjustments)?,
        "flowChartConnector" => preset_geometry_to_path("ellipse", adjustments)?,
        "flowChartInputOutput" | "flowChartManualInput" => {
            preset_geometry_to_path("parallelogram", adjustments)?
        }
        "flowChartTerminator" => preset_geometry_to_path("roundRect", adjustments)?,
        _ => return None,
    };
    Some(result)
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

fn star(points: usize) -> Vec<GeometryPathCommand> {
    polygon(
        &(0..points * 2)
            .map(|i| {
                let a =
                    -std::f64::consts::PI / 2.0 + i as f64 * std::f64::consts::PI / points as f64;
                let r = if i % 2 == 0 { 0.5 } else { 0.225 };
                (0.5 + a.cos() * r, 0.5 + a.sin() * r)
            })
            .collect::<Vec<_>>(),
    )
}

fn clamp_fraction(value: Option<f64>, fallback: f64) -> f64 {
    value
        .filter(|v| v.is_finite())
        .map(|v| if v > 1.0 { v / 100_000.0 } else { v })
        .unwrap_or(fallback)
        .clamp(0.0, 1.0)
}

fn arrow(direction: &str, adjustment: Option<f64>) -> Vec<GeometryPathCommand> {
    let h = clamp_fraction(adjustment, 0.5);
    polygon(&[
        (0.0, 0.25),
        (1.0 - h, 0.25),
        (1.0 - h, 0.0),
        (1.0, 0.5),
        (1.0 - h, 1.0),
        (1.0 - h, 0.75),
        (0.0, 0.75),
    ])
    .into_iter()
    .map(|c| match c {
        GeometryPathCommand::Move { x, y } => {
            let (x, y) = orient(direction, x, y);
            GeometryPathCommand::Move { x, y }
        }
        GeometryPathCommand::Line { x, y } => {
            let (x, y) = orient(direction, x, y);
            GeometryPathCommand::Line { x, y }
        }
        c => c,
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
    let b = clamp_fraction(adjustment, 0.5);
    if segments <= 2 {
        return vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: b, y: 0.0 },
            GeometryPathCommand::Line { x: b, y: 1.0 },
            GeometryPathCommand::Line { x: 1.0, y: 1.0 },
        ];
    };
    let mut c = vec![GeometryPathCommand::Move { x: 0.0, y: 0.0 }];
    for i in 1..segments {
        let f = i as f64 / segments as f64;
        let (x, y) = if i % 2 == 1 {
            (if i == 1 { b } else { f }, (i - 1) as f64 / segments as f64)
        } else {
            ((i - 1) as f64 / segments as f64, f)
        };
        c.push(GeometryPathCommand::Line { x, y })
    }
    c.push(GeometryPathCommand::Line { x: 1.0, y: 1.0 });
    c
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
    };
    let mut c = vec![GeometryPathCommand::Move { x: 0.0, y: 0.0 }];
    for i in 0..segments - 1 {
        let s = i as f64 / (segments - 1) as f64;
        let e = (i + 1) as f64 / (segments - 1) as f64;
        c.push(GeometryPathCommand::Cubic {
            cp1x: s + (e - s) * 0.5,
            cp1y: s,
            cp2x: s + (e - s) * 0.5,
            cp2y: e,
            x: e,
            y: e,
        })
    }
    c
}

pub fn parse_custom_geometry_path(
    custom_geometry: Option<&XmlElement>,
) -> Option<Vec<GeometryPathCommand>> {
    let path_list = local_child(custom_geometry, "pathLst")?;
    let mut result = Vec::new();
    for path in path_list.children_by_local_name("path") {
        if result.len() >= MAX_CUSTOM_PATH_COMMANDS {
            break;
        }
        let width = finite_path_attr(Some(path), "w").unwrap_or(1.0);
        let height = finite_path_attr(Some(path), "h").unwrap_or(1.0);
        let values = build_custom_guides(custom_geometry, width, height);
        result.extend(parse_custom_path(
            path,
            MAX_CUSTOM_PATH_COMMANDS - result.len(),
            &values,
        ));
    }
    (!result.is_empty()).then_some(result)
}

fn parse_custom_path(
    path: &XmlElement,
    remaining: usize,
    geometry_values: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let path_width = guide_path_attr(Some(path), "w", geometry_values);
    let path_height = guide_path_attr(Some(path), "h", geometry_values);
    let mut values = geometry_values.clone();
    if let Some(width) = path_width.filter(|value| *value > 0.0) {
        values.insert("w".into(), width);
        values.insert("r".into(), width);
        values.insert("hc".into(), width / 2.0);
    }
    if let Some(height) = path_height.filter(|value| *value > 0.0) {
        values.insert("h".into(), height);
        values.insert("b".into(), height);
        values.insert("vc".into(), height / 2.0);
    }

    let mut raw = Vec::new();
    let mut current = None;
    for child in path.child_elements() {
        if raw.len() >= remaining {
            break;
        }
        match child.local_name() {
            "moveTo" | "lnTo" => {
                let Some(point) = path_point(local_child(Some(child), "pt"), &values) else {
                    continue;
                };
                raw.push(if child.local_name() == "moveTo" {
                    GeometryPathCommand::Move {
                        x: point.0,
                        y: point.1,
                    }
                } else {
                    GeometryPathCommand::Line {
                        x: point.0,
                        y: point.1,
                    }
                });
                current = Some(point);
            }
            "quadBezTo" => {
                let points = path_points(child, &values);
                if points.len() >= 2 {
                    raw.push(GeometryPathCommand::Quad {
                        cpx: points[0].0,
                        cpy: points[0].1,
                        x: points[1].0,
                        y: points[1].1,
                    });
                    current = Some(points[1]);
                }
            }
            "cubicBezTo" => {
                let points = path_points(child, &values);
                if points.len() >= 3 {
                    raw.push(GeometryPathCommand::Cubic {
                        cp1x: points[0].0,
                        cp1y: points[0].1,
                        cp2x: points[1].0,
                        cp2y: points[1].1,
                        x: points[2].0,
                        y: points[2].1,
                    });
                    current = Some(points[2]);
                }
            }
            "arcTo" => {
                let Some(point) = current else { continue };
                let Some(wr) = guide_path_attr(Some(child), "wR", &values) else {
                    continue;
                };
                let Some(hr) = guide_path_attr(Some(child), "hR", &values) else {
                    continue;
                };
                let Some(start) = guide_path_attr(Some(child), "stAng", &values) else {
                    continue;
                };
                let Some(sweep) = guide_path_attr(Some(child), "swAng", &values) else {
                    continue;
                };
                let cubics = arc_to_cubics(point, wr, hr, start, sweep);
                for command in cubics.into_iter().take(remaining - raw.len()) {
                    if let GeometryPathCommand::Cubic { x, y, .. } = &command {
                        current = Some((*x, *y));
                    }
                    raw.push(command);
                }
            }
            "close" => raw.push(GeometryPathCommand::Close),
            _ => {}
        }
    }
    normalize_raw_path(raw, path_width, path_height)
}

fn arc_to_cubics(
    current: (f64, f64),
    width_radius: f64,
    height_radius: f64,
    start_angle: f64,
    sweep_angle: f64,
) -> Vec<GeometryPathCommand> {
    let rx = width_radius.abs();
    let ry = height_radius.abs();
    if rx <= 0.0 || ry <= 0.0 || sweep_angle == 0.0 {
        return Vec::new();
    }
    let start = angle_radians(start_angle);
    let sweep = angle_radians(sweep_angle);
    let center = (current.0 - rx * start.cos(), current.1 - ry * start.sin());
    let count = ((sweep.abs() / (std::f64::consts::PI / 2.0)).ceil() as usize).max(1);
    let segment_sweep = sweep / count as f64;
    let mut output = Vec::with_capacity(count);
    let mut p0 = current;
    for index in 0..count {
        let t0 = start + segment_sweep * index as f64;
        let t1 = t0 + segment_sweep;
        let alpha = 4.0 / 3.0 * (segment_sweep / 4.0).tan();
        let p1 = (center.0 + rx * t1.cos(), center.1 + ry * t1.sin());
        let d0 = (-rx * t0.sin(), ry * t0.cos());
        let d1 = (-rx * t1.sin(), ry * t1.cos());
        output.push(GeometryPathCommand::Cubic {
            cp1x: p0.0 + alpha * d0.0,
            cp1y: p0.1 + alpha * d0.1,
            cp2x: p1.0 - alpha * d1.0,
            cp2y: p1.1 - alpha * d1.1,
            x: p1.0,
            y: p1.1,
        });
        p0 = p1;
    }
    output
}

fn normalize_raw_path(
    commands: Vec<GeometryPathCommand>,
    path_width: Option<f64>,
    path_height: Option<f64>,
) -> Vec<GeometryPathCommand> {
    let points = commands.iter().flat_map(command_points).collect::<Vec<_>>();
    let width = path_width
        .filter(|value| *value > 0.0)
        .unwrap_or_else(|| positive_extent(&points, true));
    let height = path_height
        .filter(|value| *value > 0.0)
        .unwrap_or_else(|| positive_extent(&points, false));
    commands
        .into_iter()
        .map(|command| normalize_command(command, width, height))
        .collect()
}

fn command_points(command: &GeometryPathCommand) -> Vec<(f64, f64)> {
    match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => vec![(*x, *y)],
        GeometryPathCommand::Quad { cpx, cpy, x, y } => vec![(*cpx, *cpy), (*x, *y)],
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            vec![(*cp1x, *cp1y), (*cp2x, *cp2y), (*x, *y)]
        }
        GeometryPathCommand::Close => Vec::new(),
    }
}

fn positive_extent(points: &[(f64, f64)], x_axis: bool) -> f64 {
    let max = points.iter().fold(0.0_f64, |value, point| {
        value.max(if x_axis { point.0 } else { point.1 })
    });
    if max > 0.0 { max } else { 1.0 }
}

fn normalize_command(command: GeometryPathCommand, width: f64, height: f64) -> GeometryPathCommand {
    let x = |value| normalize_path_number(value, width);
    let y = |value| normalize_path_number(value, height);
    match command {
        GeometryPathCommand::Move { x: px, y: py } => {
            GeometryPathCommand::Move { x: x(px), y: y(py) }
        }
        GeometryPathCommand::Line { x: px, y: py } => {
            GeometryPathCommand::Line { x: x(px), y: y(py) }
        }
        GeometryPathCommand::Quad {
            cpx,
            cpy,
            x: px,
            y: py,
        } => GeometryPathCommand::Quad {
            cpx: x(cpx),
            cpy: y(cpy),
            x: x(px),
            y: y(py),
        },
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x: px,
            y: py,
        } => GeometryPathCommand::Cubic {
            cp1x: x(cp1x),
            cp1y: y(cp1y),
            cp2x: x(cp2x),
            cp2y: y(cp2y),
            x: x(px),
            y: y(py),
        },
        GeometryPathCommand::Close => GeometryPathCommand::Close,
    }
}

fn normalize_path_number(value: f64, denominator: f64) -> f64 {
    let value = value / denominator;
    if value.is_finite() {
        value.clamp(
            -MAX_ABS_NORMALIZED_PATH_NUMBER,
            MAX_ABS_NORMALIZED_PATH_NUMBER,
        )
    } else {
        0.0
    }
}

fn build_custom_guides(
    geometry: Option<&XmlElement>,
    width: f64,
    height: f64,
) -> HashMap<String, f64> {
    let mut values = standard_guide_values(width, height);
    apply_guide_list(geometry, "avLst", &mut values);
    apply_guide_list(geometry, "gdLst", &mut values);
    values
}

fn standard_guide_values(width: f64, height: f64) -> HashMap<String, f64> {
    let width = if width > 0.0 { width } else { 1.0 };
    let height = if height > 0.0 { height } else { 1.0 };
    let short = width.min(height);
    let long = width.max(height);
    let mut values = HashMap::from([
        ("l".into(), 0.0),
        ("t".into(), 0.0),
        ("r".into(), width),
        ("b".into(), height),
        ("w".into(), width),
        ("h".into(), height),
        ("hc".into(), width / 2.0),
        ("vc".into(), height / 2.0),
        ("ss".into(), short),
        ("ls".into(), long),
        ("cd2".into(), 10_800_000.0),
        ("cd4".into(), 5_400_000.0),
        ("cd8".into(), 2_700_000.0),
        ("3cd4".into(), 16_200_000.0),
        ("3cd8".into(), 8_100_000.0),
        ("5cd8".into(), 13_500_000.0),
        ("7cd8".into(), 18_900_000.0),
    ]);
    for divisor in [2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 32] {
        let divisor_f64 = divisor as f64;
        values.insert(format!("wd{divisor}"), width / divisor_f64);
        values.insert(format!("hd{divisor}"), height / divisor_f64);
        values.insert(format!("ssd{divisor}"), short / divisor_f64);
        values.insert(format!("lsd{divisor}"), long / divisor_f64);
    }
    values
}

fn apply_guide_list(
    parent: Option<&XmlElement>,
    list_name: &str,
    values: &mut HashMap<String, f64>,
) {
    let Some(list) = local_child(parent, list_name) else {
        return;
    };
    let mut pending = list
        .children_by_local_name("gd")
        .take(MAX_CUSTOM_GUIDES)
        .collect::<Vec<_>>();
    for _ in 0..pending.len() {
        if pending.is_empty() {
            break;
        }
        for index in (0..pending.len()).rev() {
            let guide = pending[index];
            let Some(name) = guide.attribute(None, "name") else {
                continue;
            };
            let Some(value) = evaluate_guide(guide.attribute(None, "fmla"), values) else {
                continue;
            };
            values.insert(name.to_owned(), value);
            pending.remove(index);
        }
    }
}

fn evaluate_guide(formula: Option<&str>, values: &HashMap<String, f64>) -> Option<f64> {
    let mut tokens = formula?.trim().split_whitespace();
    let operator = tokens.next()?;
    let args = tokens
        .map(|token| guide_token(token, values))
        .collect::<Option<Vec<_>>>()?;
    let x = args.first().copied().unwrap_or(0.0);
    let y = args.get(1).copied().unwrap_or(0.0);
    let z = args.get(2).copied().unwrap_or(0.0);
    let result = match operator {
        "val" => x,
        "*/" if z != 0.0 => x * y / z,
        "+-" => x + y - z,
        "+/" if z != 0.0 => (x + y) / z,
        "?:" => {
            if x > 0.0 {
                y
            } else {
                z
            }
        }
        "abs" => x.abs(),
        "at2" => y.atan2(x) * 180.0 * ANGLE_UNITS_PER_DEGREE / std::f64::consts::PI,
        "cat2" => x * z.atan2(y).cos(),
        "cos" => x * angle_radians(y).cos(),
        "max" => x.max(y),
        "min" => x.min(y),
        "mod" => (x * x + y * y + z * z).sqrt(),
        "pin" => y.max(x).min(z),
        "sat2" => x * z.atan2(y).sin(),
        "sin" => x * angle_radians(y).sin(),
        "sqrt" if x >= 0.0 => x.sqrt(),
        "tan" => x * angle_radians(y).tan(),
        _ => return None,
    };
    (result.is_finite() && result.abs() <= MAX_ABS_CUSTOM_PATH_NUMBER).then_some(result)
}

fn guide_token(token: &str, values: &HashMap<String, f64>) -> Option<f64> {
    finite_path_number(token).or_else(|| values.get(token).copied())
}

fn finite_path_number(value: &str) -> Option<f64> {
    let value = if value.trim().is_empty() {
        0.0
    } else {
        value.trim().parse::<f64>().ok()?
    };
    (value.is_finite() && value.abs() <= MAX_ABS_CUSTOM_PATH_NUMBER).then_some(value)
}

fn finite_path_attr(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    finite_path_number(element?.attribute(None, name)?)
}

fn guide_path_attr(
    element: Option<&XmlElement>,
    name: &str,
    values: &HashMap<String, f64>,
) -> Option<f64> {
    guide_token(element?.attribute(None, name)?, values)
}

fn path_point(element: Option<&XmlElement>, values: &HashMap<String, f64>) -> Option<(f64, f64)> {
    Some((
        guide_path_attr(element, "x", values)?,
        guide_path_attr(element, "y", values)?,
    ))
}

fn path_points(command: &XmlElement, values: &HashMap<String, f64>) -> Vec<(f64, f64)> {
    command
        .children_by_local_name("pt")
        .filter_map(|point| path_point(Some(point), values))
        .collect()
}

fn angle_radians(value: f64) -> f64 {
    value / ANGLE_UNITS_PER_DEGREE * (std::f64::consts::PI / 180.0)
}

fn numeric_attr(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    parse_javascript_integer_prefix(element?.attribute(None, name)?)
}
fn local_child<'a>(parent: Option<&'a XmlElement>, name: &str) -> Option<&'a XmlElement> {
    parent?
        .child_elements()
        .find(|child| child.local_name() == name)
}
fn js_round(value: f64) -> f64 {
    (value + 0.5).floor()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(
            xml.as_bytes(),
            "drawing.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    #[test]
    fn colors_fills_outlines_and_null_line_ends_match_the_shared_package() {
        let solid = root(
            "<a:solidFill><a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr></a:solidFill>",
        );
        assert_eq!(
            parse_color_element(Some(&solid)),
            Some(ColorValue {
                theme_color: Some("accent1".into()),
                theme_tint: Some("66".into()),
                ..ColorValue::default()
            })
        );
        let properties = root(
            "<wps:spPr><a:gradFill/><a:ln w=\"19050px\"><a:solidFill><a:srgbClr val=\"00B0F0\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></wps:spPr>",
        );
        assert_eq!(parse_fill(Some(&properties)).unwrap().fill_type, "gradient");
        assert_eq!(
            parse_outline(Some(&properties)).unwrap().width,
            Some(19_050.0)
        );
        let tail = parse_line_end(&root("<a:tailEnd type=\"oval\"/>"));
        assert_eq!(
            serde_json::to_value(tail).unwrap(),
            serde_json::json!({"type":"oval","width":null,"length":null})
        );
    }

    #[test]
    fn preset_and_custom_geometry_are_bounded_and_normalized() {
        let properties =
            root("<wps:spPr><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom></wps:spPr>");
        assert_eq!(
            parse_preset_geometry_path(Some(&properties)).unwrap().len(),
            10
        );
        let geometry = root(
            "<a:custGeom><a:gdLst><a:gd name=\"x1\" fmla=\"*/ w 1 2\"/></a:gdLst><a:pathLst><a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"x1\" y=\"100\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>",
        );
        assert_eq!(
            parse_custom_geometry_path(Some(&geometry)).unwrap(),
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 0.5, y: 1.0 },
                GeometryPathCommand::Close
            ]
        );
        let hostile = root(
            "<a:custGeom><a:pathLst><a:path w=\"1\" h=\"1\"><a:moveTo><a:pt x=\"1e999\" y=\"0\"/></a:moveTo></a:path></a:pathLst></a:custGeom>",
        );
        assert_eq!(parse_custom_geometry_path(Some(&hostile)), None);
    }

    #[test]
    fn transforms_and_gradients_keep_javascript_integer_prefixes() {
        let transform = root(
            "<a:xfrm rot=\"2700000junk\" flipH=\"1\"><a:off x=\"10\" y=\"20\"/><a:ext cx=\"1828800\" cy=\"914400\"/></a:xfrm>",
        );
        let parsed = parse_transform(Some(&transform));
        assert_eq!(parsed.transform.unwrap().rotation, Some(45.0));
        assert_eq!(parsed.offset, Some(Point2D { x: 10.0, y: 20.0 }));
        let gradient = root(
            "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs></a:gsLst><a:lin ang=\"5400000\"/></a:gradFill>",
        );
        assert_eq!(
            parse_gradient_fill(&gradient).gradient.unwrap().angle,
            Some(90.0)
        );
    }
}
