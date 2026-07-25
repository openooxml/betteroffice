use std::collections::BTreeMap;

use ooxml_drawingml::{ColorValue, ShapeFill, ShapeOutline, Theme};
use serde::{Deserialize, Serialize};

use crate::relationships::Relationship;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PptxPackage {
    pub presentation: Presentation,
    pub slides: Vec<Slide>,
    pub layouts: Vec<SlideLayout>,
    pub masters: Vec<SlideMaster>,
    pub themes: Vec<ThemePart>,
    pub media: Vec<MediaPart>,
    pub relationships: BTreeMap<String, Vec<Relationship>>,
    #[serde(skip)]
    pub(crate) parts: Vec<PackagePart>,
}

impl PptxPackage {
    pub fn part_bytes(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|part| part.path == path)
            .map(|part| part.bytes.as_slice())
    }

    pub fn replace_part(&mut self, path: &str, bytes: Vec<u8>) -> bool {
        let Some(part) = self.parts.iter_mut().find(|part| part.path == path) else {
            return false;
        };
        part.bytes = bytes;
        true
    }
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct PackagePart {
    pub path: String,
    pub bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Presentation {
    pub part_path: String,
    pub width_emu: i64,
    pub height_emu: i64,
    pub slides: Vec<SlideReference>,
    pub master_part_paths: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideReference {
    pub id: u32,
    pub relationship_id: String,
    pub part_path: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Slide {
    pub part_path: String,
    pub name: Option<String>,
    pub layout_part_path: Option<String>,
    pub show_master_shapes: bool,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideLayout {
    pub part_path: String,
    pub name: Option<String>,
    pub layout_type: Option<String>,
    pub master_part_path: Option<String>,
    pub show_master_shapes: bool,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideMaster {
    pub part_path: String,
    pub name: Option<String>,
    pub theme_part_path: Option<String>,
    pub layout_part_paths: Vec<String>,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
    pub text_styles: TextStyleSet,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemePart {
    pub part_path: String,
    pub theme: Theme,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaPart {
    pub part_path: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum ShapeNode {
    Shape(Shape),
    Picture(Picture),
    GraphicFrame(GraphicFrame),
    Group(GroupShape),
}

impl ShapeNode {
    pub fn id(&self) -> u32 {
        match self {
            Self::Shape(shape) => shape.base.id,
            Self::Picture(picture) => picture.base.id,
            Self::GraphicFrame(frame) => frame.base.id,
            Self::Group(group) => group.base.id,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeBase {
    pub id: u32,
    pub name: String,
    pub description: Option<String>,
    pub hidden: bool,
    pub placeholder: Option<Placeholder>,
    pub transform: ShapeTransform,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTransform {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub child_x: Option<i64>,
    pub child_y: Option<i64>,
    pub child_width: Option<i64>,
    pub child_height: Option<i64>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Placeholder {
    pub placeholder_type: Option<String>,
    pub index: Option<u32>,
    pub orientation: Option<String>,
    pub size: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Shape {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub geometry: String,
    #[serde(default)]
    pub adjust_values: BTreeMap<String, f64>,
    pub fill: Option<ShapeFill>,
    pub outline: Option<ShapeOutline>,
    pub text: Option<TextBody>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Picture {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub relationship_id: Option<String>,
    pub media_part_path: Option<String>,
    pub crop: PictureCrop,
    pub fill: Option<ShapeFill>,
    pub outline: Option<ShapeOutline>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct PictureCrop {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GraphicFrame {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub data: GraphicFrameData,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum GraphicFrameData {
    Table {
        rows: Vec<Vec<TextBody>>,
    },
    Chart {
        relationship_id: String,
        part_path: Option<String>,
    },
    Diagram {
        relationship_ids: Vec<String>,
    },
    Unknown {
        uri: Option<String>,
    },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GroupShape {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub children: Vec<ShapeNode>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBody {
    pub anchor: Option<String>,
    pub vertical: Option<String>,
    pub autofit: Option<TextAutofit>,
    pub inset_left: Option<i64>,
    pub inset_top: Option<i64>,
    pub inset_right: Option<i64>,
    pub inset_bottom: Option<i64>,
    pub paragraphs: Vec<TextParagraph>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum TextAutofit {
    None,
    Shape,
    Normal {
        font_scale: Option<f64>,
        line_space_reduction: Option<f64>,
    },
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextParagraph {
    pub properties: ParagraphProperties,
    pub runs: Vec<TextRun>,
    pub end_properties: Option<RunProperties>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphProperties {
    pub alignment: Option<String>,
    pub level: u32,
    pub margin_left: Option<i64>,
    pub indent: Option<i64>,
    pub bullet: Option<Bullet>,
    pub default_run: Option<RunProperties>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum Bullet {
    Character { value: String },
    AutoNumber { scheme: String, start_at: u32 },
    None,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub properties: RunProperties,
    pub field_id: Option<String>,
    pub field_type: Option<String>,
    pub line_break: bool,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunProperties {
    pub font_size_pt: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub font_family: Option<String>,
    pub color: Option<ColorValue>,
    pub language: Option<String>,
    pub hyperlink_relationship_id: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStyleSet {
    pub title: Vec<ParagraphProperties>,
    pub body: Vec<ParagraphProperties>,
    pub other: Vec<ParagraphProperties>,
}
