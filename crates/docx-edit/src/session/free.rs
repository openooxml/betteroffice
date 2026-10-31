use serde::Serialize;
use serde_json::json;

use super::{EngineError, EngineResult};
use crate::batch::outcome_json;
use crate::content_controls::{ContentControlQuery, ContentControlsOptions};
use crate::structured::ExportOptions;

pub fn snapshot_json<T: Serialize>(
    outcome: Result<T, crate::structured::ExportError>,
) -> EngineResult<String> {
    let outcome = match outcome {
        Ok(content) => Ok(json!({ "content": content })),
        Err(crate::structured::ExportError::Refused(failure)) => Err(json!({ "failure": failure })),
        Err(crate::structured::ExportError::Parse(message)) => {
            return Err(EngineError::message(message));
        }
    };
    outcome_json(&outcome).map_err(EngineError::message)
}

pub fn export_docx_structured_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    let options: ExportOptions = serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(crate::structured::export_docx_structured(bytes, &options))
}

pub fn export_docx_markdown_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    let options: ExportOptions = serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(crate::structured::export_docx_markdown(bytes, &options))
}

pub fn list_docx_content_controls_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    let options: ContentControlsOptions =
        serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(crate::content_controls::list_docx_content_controls(
        bytes, &options,
    ))
}

pub fn find_docx_content_controls_json(
    bytes: &[u8],
    query: &str,
    options: &str,
) -> EngineResult<String> {
    let query: ContentControlQuery = serde_json::from_str(query).map_err(EngineError::message)?;
    let options: ContentControlsOptions =
        serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(crate::content_controls::find_docx_content_controls(
        bytes, &query, &options,
    ))
}

pub fn render_docx_markdown_json(content: &str, options: &str) -> EngineResult<String> {
    let content: crate::structured::DocxStructuredContent =
        serde_json::from_str(content).map_err(EngineError::message)?;
    let options: crate::structured::MarkdownOptions =
        serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(
        crate::structured::render_docx_markdown(&content, &options)
            .map_err(crate::structured::ExportError::Refused),
    )
}

pub fn render_docx_markdown_with_pages_json(content: &str, options: &str) -> EngineResult<String> {
    #[derive(serde::Deserialize)]
    #[serde(rename_all = "camelCase")]
    struct Layout {
        #[serde(
            rename = "schemaVersion",
            deserialize_with = "crate::structured::pages::page_map_schema_version"
        )]
        _schema_version: u8,
        export_fingerprint: String,
        pages: Vec<crate::structured::ExportPage>,
        fragments: Vec<crate::structured::PageFragment>,
    }
    #[derive(serde::Deserialize)]
    struct Paged {
        structured: crate::structured::DocxStructuredContent,
        layout: Layout,
    }
    let content: Paged = serde_json::from_str(content).map_err(EngineError::message)?;
    let options: crate::structured::PageMarkdownOptions =
        serde_json::from_str(options).map_err(EngineError::message)?;
    snapshot_json(
        crate::structured::render_docx_markdown_with_pages(
            &content.structured,
            crate::structured::PageAnnotations {
                export_fingerprint: &content.layout.export_fingerprint,
                pages: &content.layout.pages,
                fragments: &content.layout.fragments,
            },
            &options,
        )
        .map_err(crate::structured::ExportError::Refused),
    )
}
