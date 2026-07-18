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
