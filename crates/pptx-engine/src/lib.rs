//! Target-neutral PPTX engine boundary shared by the wasm and native backends.

mod float_bits;
mod renderer;

use serde::Serialize;
use std::collections::BTreeMap;
use std::fmt;

use pptx_edit::structured;

pub use pptx_edit::engine::{Document, EngineError, ErrorKind, Host, version};
pub use renderer::Renderer;

pub fn renderer_version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

pub(crate) fn engine_error(error: impl fmt::Display) -> EngineError {
    EngineError::message(error.to_string())
}

pub fn parse_pptx_json(data: &[u8]) -> Result<String, EngineError> {
    let package = pptx_parse::parse_pptx(data).map_err(engine_error)?;
    serde_json::to_string(&package).map_err(engine_error)
}

pub fn parse_pptx_json_without_media(data: &[u8]) -> Result<String, EngineError> {
    let package = pptx_parse::parse_pptx(data).map_err(engine_error)?;
    let metadata = PptxPackageWithoutMedia {
        presentation: &package.presentation,
        slides: &package.slides,
        layouts: &package.layouts,
        masters: &package.masters,
        themes: &package.themes,
        charts: &package.charts,
        diagram_drawings: &package.diagram_drawings,
        media: package
            .media
            .iter()
            .map(|part| MediaSummary {
                part_path: &part.part_path,
                content_type: &part.content_type,
                byte_length: part.bytes.len(),
            })
            .collect(),
        table_styles: &package.table_styles,
        comment_authors: &package.comment_authors,
        comments: &package.comments,
        comment_flavor: &package.comment_flavor,
        relationships: &package.relationships,
        shape_elements: package.models_connectors().then_some("withConnectors"),
    };
    serde_json::to_string(&metadata).map_err(engine_error)
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct PptxPackageWithoutMedia<'a> {
    presentation: &'a pptx_parse::Presentation,
    slides: &'a [pptx_parse::Slide],
    layouts: &'a [pptx_parse::SlideLayout],
    masters: &'a [pptx_parse::SlideMaster],
    themes: &'a [pptx_parse::ThemePart],
    charts: &'a [pptx_parse::ChartPart],
    #[serde(skip_serializing_if = "Vec::is_empty")]
    diagram_drawings: &'a Vec<pptx_parse::DiagramDrawing>,
    media: Vec<MediaSummary<'a>>,
    #[serde(skip_serializing_if = "ooxml_drawingml::TableStyleList::is_empty")]
    table_styles: &'a ooxml_drawingml::TableStyleList,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    comment_authors: &'a Vec<pptx_parse::CommentAuthor>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    comments: &'a Vec<pptx_parse::Comment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    comment_flavor: &'a Option<pptx_parse::CommentFlavor>,
    relationships: &'a BTreeMap<String, Vec<pptx_parse::Relationship>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    shape_elements: Option<&'static str>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct MediaSummary<'a> {
    part_path: &'a str,
    content_type: &'a str,
    byte_length: usize,
}

pub fn export_pptx_structured_json(data: &[u8], options: &str) -> Result<String, EngineError> {
    let options = serde_json::from_str(options).map_err(engine_error)?;
    structured::snapshot_outcome_json(structured::export_pptx_structured(data, &options))
        .map_err(EngineError::message)
}

pub fn export_pptx_markdown_json(data: &[u8], options: &str) -> Result<String, EngineError> {
    let options = serde_json::from_str(options).map_err(engine_error)?;
    structured::snapshot_outcome_json(structured::export_pptx_markdown(data, &options))
        .map_err(EngineError::message)
}

pub fn render_pptx_markdown_json(content: &str, options: &str) -> Result<String, EngineError> {
    let content: structured::PptxStructuredContent =
        serde_json::from_str(content).map_err(engine_error)?;
    let options: structured::PptxMarkdownOptions =
        serde_json::from_str(options).map_err(engine_error)?;
    structured::snapshot_outcome_json(
        structured::render_pptx_markdown(&content, &options)
            .map_err(structured::ExportError::Refused),
    )
    .map_err(EngineError::message)
}

pub fn compile_slide_json(slide_json: &str) -> Result<String, EngineError> {
    pptx_render::compile_json(slide_json).map_err(EngineError::message)
}

pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, EngineError> {
    ooxml_drawingml::media::decode_tiff_png(data).map_err(engine_error)
}
