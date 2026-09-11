use serde::{Deserialize, Serialize};

use crate::ColorValue;

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
    pub fn named(fill_type: &str) -> Self {
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFill>,
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

/// Shape effects.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeEffects {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outer_shadow: Option<OuterShadow>,
}

/// Outer shadow in EMUs and 60000ths of a degree.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OuterShadow {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(default, skip_serializing_if = "is_zero_emu")]
    pub blur_radius: i64,
    #[serde(default, skip_serializing_if = "is_zero_emu")]
    pub distance: i64,
    #[serde(default, skip_serializing_if = "is_zero_emu")]
    pub direction: i64,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub rotate_with_shape: bool,
    /// `sx`/`sy`: the shadow's own size, as a fraction of the shape's.
    #[serde(default = "default_scale", skip_serializing_if = "is_unit_scale")]
    pub scale_x: f64,
    #[serde(default = "default_scale", skip_serializing_if = "is_unit_scale")]
    pub scale_y: f64,
    /// `algn`: the point of the shape's box a scaled shadow keeps.
    #[serde(
        default = "default_alignment",
        skip_serializing_if = "is_default_alignment"
    )]
    pub alignment: String,
}

impl Default for OuterShadow {
    fn default() -> Self {
        Self {
            color: None,
            blur_radius: 0,
            distance: 0,
            direction: 0,
            rotate_with_shape: true,
            scale_x: 1.0,
            scale_y: 1.0,
            alignment: default_alignment(),
        }
    }
}

fn default_scale() -> f64 {
    1.0
}
fn is_unit_scale(value: &f64) -> bool {
    *value == 1.0
}
fn default_alignment() -> String {
    "b".to_owned()
}
fn is_default_alignment(value: &str) -> bool {
    value == "b"
}

fn is_zero_emu(value: &i64) -> bool {
    *value == 0
}
fn default_true() -> bool {
    true
}
fn is_true(value: &bool) -> bool {
    *value
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct LineEnd {
    #[serde(rename = "type")]
    pub end_type: String,
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

pub fn resolve_shape_fill_color(fill_type: Option<&str>, color: Option<&str>) -> Option<String> {
    if fill_type == Some("none") {
        return None;
    }
    Some(color.unwrap_or("#ffffff").to_owned())
}
