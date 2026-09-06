use ooxml_drawingml::GeometryPathCommand;
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end: Option<StrokeEnd>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<StrokeEnd>,
}

/// End dimensions in CSS pixels.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StrokeEnd {
    pub kind: String,
    pub width: f32,
    pub length: f32,
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

/// Source fractions discarded by `a:srcRect`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageCrop {
    #[serde(default, skip_serializing_if = "is_zero")]
    pub left: f32,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub top: f32,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub right: f32,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub bottom: f32,
}

impl ImageCrop {
    pub fn is_whole(&self) -> bool {
        *self == Self::default()
    }

    /// The fraction of the source that survives on each axis.
    pub fn kept(&self) -> (f32, f32) {
        (1.0 - self.left - self.right, 1.0 - self.top - self.bottom)
    }
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
        #[serde(default, skip_serializing_if = "Option::is_none")]
        shape_id: Option<String>,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        geometry: String,
        path: Vec<GeometryPathCommand>,
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
        #[serde(default, skip_serializing_if = "Option::is_none")]
        shape_id: Option<String>,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        asset_id: Option<String>,
        #[serde(default, skip_serializing_if = "ImageCrop::is_whole")]
        crop: ImageCrop,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        path: Option<Vec<GeometryPathCommand>>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        stroke: Option<Stroke>,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
    TextBox {
        object_id: u32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        shape_id: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        story_id: Option<String>,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        anchor: TextAnchor,
        paragraphs: Vec<TextParagraph>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        lines: Vec<PositionedTextLine>,
        #[serde(default, skip_serializing_if = "is_false")]
        overflow: bool,
        #[serde(default, skip_serializing_if = "Transform::is_identity")]
        transform: Transform,
    },
    Placeholder {
        object_id: u32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        shape_id: Option<String>,
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
    /// A plotted chart: one addressable object whose parts paint clipped to
    /// its rectangle, and whose `label` is the screen-reader summary.
    Chart {
        object_id: u32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        shape_id: Option<String>,
        name: String,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        label: String,
        primitives: Vec<Primitive>,
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

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionedTextLine {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub baseline: f32,
    pub start: u32,
    pub end: u32,
    pub runs: Vec<PositionedTextRun>,
    pub caret_stops: Vec<CaretStop>,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaretStop {
    pub position: u32,
    pub x: f32,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionedTextRun {
    pub text: String,
    pub start: u32,
    pub end: u32,
    pub x: f32,
    pub width: f32,
    pub font_id: u32,
    pub font_family: String,
    pub font_size_px: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: String,
    pub glyphs: Vec<PositionedGlyph>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionedGlyph {
    pub glyph_id: u32,
    pub cluster: u32,
    pub x: f32,
    pub advance: f32,
    pub x_offset: f32,
    pub y_offset: f32,
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
                shape_id: None,
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

    #[test]
    fn an_uncropped_rectangular_image_serializes_as_it_did_before_crops_existed() {
        let mut image = Primitive::Image {
            object_id: 90,
            shape_id: Some("slide:0:256:shape:9".into()),
            name: "Media fixture".into(),
            x: 1280.0,
            y: 720.0,
            w: 0.5,
            h: 0.25,
            asset_id: Some("ppt/media/betteroffice-mark.png".into()),
            crop: ImageCrop::default(),
            path: None,
            stroke: None,
            transform: Transform::default(),
        };
        let before = r#"{"kind":"image","objectId":90,"shapeId":"slide:0:256:shape:9","name":"Media fixture","x":1280.0,"y":720.0,"w":0.5,"h":0.25,"assetId":"ppt/media/betteroffice-mark.png"}"#;
        assert_eq!(serde_json::to_string(&image).unwrap(), before);
        assert_eq!(serde_json::from_str::<Primitive>(before).unwrap(), image);

        let Primitive::Image { crop, path, .. } = &mut image else {
            unreachable!()
        };
        *crop = ImageCrop {
            top: 0.1,
            bottom: 0.2,
            ..ImageCrop::default()
        };
        *path = Some(vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.5 },
            GeometryPathCommand::Close,
        ]);
        let json = serde_json::to_string(&image).unwrap();
        assert!(json.contains(r#""crop":{"top":0.1,"bottom":0.2}"#));
        assert!(json.contains(r#""path":[{"type":"move","x":0.0,"y":0.5},{"type":"close"}]"#));
        assert_eq!(serde_json::from_str::<Primitive>(&json).unwrap(), image);
    }
}
