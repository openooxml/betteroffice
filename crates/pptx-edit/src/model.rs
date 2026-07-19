use ooxml_drawingml::{ShapeFill, ShapeOutline};
use pptx_parse::{GraphicFrameData, Placeholder};
use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EditOrigin {
    #[default]
    Local,
    Agent,
    Remote,
    System,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EditCtx {
    pub origin: EditOrigin,
    pub author: String,
}

impl EditCtx {
    pub fn local(author: impl Into<String>) -> Self {
        Self {
            origin: EditOrigin::Local,
            author: author.into(),
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStylePatch {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRunSnapshot {
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphSnapshot {
    pub id: String,
    pub alignment: Option<String>,
    pub level: u32,
    pub bullet_json: Option<String>,
    pub runs: Vec<TextRunSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StorySnapshot {
    pub id: String,
    pub length: u32,
    pub paragraphs: Vec<ParagraphSnapshot>,
}

impl StorySnapshot {
    pub fn plain_text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeSnapshot {
    pub id: String,
    pub source_id: u32,
    pub kind: ShapeKind,
    pub name: String,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub geometry: String,
    pub placeholder: Option<Placeholder>,
    pub fill: Option<ShapeFill>,
    pub outline: Option<ShapeOutline>,
    pub media_part_path: Option<String>,
    pub graphic: Option<GraphicFrameData>,
    pub text_stories: Vec<StorySnapshot>,
    pub children: Vec<ShapeSnapshot>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShapeKind {
    Shape,
    Picture,
    GraphicFrame,
    Group,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideSnapshot {
    pub id: String,
    pub source_part_path: Option<String>,
    pub layout_part_path: Option<String>,
    pub name: Option<String>,
    pub shapes: Vec<ShapeSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DeckSnapshot {
    pub width_emu: i64,
    pub height_emu: i64,
    pub slides: Vec<SlideSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideReceipt {
    pub slide_id: String,
    pub from_index: Option<u32>,
    pub to_index: Option<u32>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub index: u32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TransformReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: ShapeRect,
    pub after: ShapeRect,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ShapeRect {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextReceipt {
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDraft {
    pub name: String,
    pub rect: ShapeRect,
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UpdateOrigin {
    Local,
    Remote,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct UpdateEvent {
    pub update: Vec<u8>,
    pub origin: UpdateOrigin,
}

#[derive(Debug, Error)]
pub enum EditError {
    #[error("invalid client ID {0}")]
    InvalidClientId(u64),
    #[error("could not parse PPTX: {0}")]
    Parse(String),
    #[error("invalid deck state: {0}")]
    InvalidState(String),
    #[error("invalid yrs update: {0}")]
    InvalidUpdate(String),
    #[error("invalid yrs state vector: {0}")]
    InvalidStateVector(String),
    #[error("slide {0:?} was not found")]
    SlideNotFound(String),
    #[error("shape {0:?} was not found")]
    ShapeNotFound(String),
    #[error("story {0:?} was not found")]
    StoryNotFound(String),
    #[error("index {index} is outside length {length}")]
    OutOfBounds { index: u32, length: u32 },
    #[error("text range {start}..{end} crosses a paragraph boundary")]
    ParagraphBoundary { start: u32, end: u32 },
    #[error("invalid shape geometry: {0}")]
    InvalidGeometry(String),
    #[error("update observer failed: {0}")]
    Observer(String),
    #[error("JSON boundary error: {0}")]
    Json(String),
}

pub type EditResult<T> = Result<T, EditError>;
