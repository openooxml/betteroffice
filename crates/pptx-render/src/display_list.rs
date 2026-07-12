use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

pub const CONTRACT_VERSION: u32 = 1;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SurfaceDisplayList {
    pub contract_version: u32,
    pub width: f32,
    pub height: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Paint>,
    pub primitives: Vec<Primitive>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum Paint {
    Solid {
        color: String,
    },
    Gradient {
        gradient_type: GradientType,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        angle_deg: Option<f32>,
        stops: Vec<GradientStop>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum GradientType {
    Linear,
    Radial,
    Rectangular,
    Path,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    pub position: f32,
    pub color: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Stroke {
    pub color: String,
    pub width: f32,
    #[serde(default, skip_serializing_if = "is_false")]
    pub dashed: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transform {
    #[serde(default, skip_serializing_if = "is_zero")]
    pub rotation_deg: f32,
    #[serde(default, skip_serializing_if = "is_false")]
    pub flip_h: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub flip_v: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum Primitive {
    Shape {
        object_id: u32,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        geometry: String,
        #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
        adjust_values: BTreeMap<String, f32>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        fill: Option<Paint>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        stroke: Option<Stroke>,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
    Image {
        object_id: u32,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        asset_id: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        stroke: Option<Stroke>,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
    TextBox {
        object_id: u32,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        anchor: TextAnchor,
        paragraphs: Vec<TextParagraph>,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
    Placeholder {
        object_id: u32,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        label: Option<String>,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextAnchor {
    Top,
    Center,
    Bottom,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextParagraph {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<TextAlign>,
    pub level: u32,
    pub runs: Vec<TextRun>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextAlign {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub font_family: String,
    pub font_size_pt: f32,
    #[serde(default, skip_serializing_if = "is_false")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub italic: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub underline: bool,
    pub color: String,
}

impl Transform {
    pub fn is_identity(&self) -> bool {
        self.rotation_deg == 0.0 && !self.flip_h && !self.flip_v
    }
}

fn is_false(value: &bool) -> bool {
    !*value
}

fn is_zero(value: &f32) -> bool {
    *value == 0.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identity_transform_is_omitted_from_json() {
        let list = SurfaceDisplayList {
            contract_version: CONTRACT_VERSION,
            width: 100.0,
            height: 50.0,
            background: None,
            primitives: vec![Primitive::Placeholder {
                object_id: 1,
                name: "chart".into(),
                x: 0.0,
                y: 0.0,
                w: 10.0,
                h: 10.0,
                label: Some("Chart".into()),
                transform: Transform::default(),
            }],
        };

        let json = serde_json::to_string(&list).expect("serialize display list");
        assert!(!json.contains("transform"));
        assert!(json.contains("contractVersion"));
    }
}
