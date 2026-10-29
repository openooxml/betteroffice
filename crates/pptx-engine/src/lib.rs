//! Target-neutral PPTX engine boundary shared by the wasm and native backends.

mod renderer;

use std::fmt;

use pptx_edit::structured;

pub use pptx_edit::engine::{
    Document, EngineError, ErrorKind, Host, parse_pptx_json, parse_pptx_json_without_media, version,
};
pub use renderer::Renderer;

pub fn renderer_version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

pub(crate) fn engine_error(error: impl fmt::Display) -> EngineError {
    EngineError::message(error.to_string())
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
