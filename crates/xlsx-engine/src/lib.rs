pub mod document;
pub mod session;

pub use document::{Document, Host, SnapshotBuilder, display_list_cell_limit};
pub use session::Session;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorKind {
    Message,
    Error,
    TypeError,
    PendingException,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EngineError {
    pub kind: ErrorKind,
    pub message: String,
}

impl EngineError {
    pub fn message(message: impl Into<String>) -> Self {
        Self {
            kind: ErrorKind::Message,
            message: message.into(),
        }
    }

    pub fn error(message: impl Into<String>) -> Self {
        Self {
            kind: ErrorKind::Error,
            message: message.into(),
        }
    }

    pub fn type_error(message: impl Into<String>) -> Self {
        Self {
            kind: ErrorKind::TypeError,
            message: message.into(),
        }
    }

    pub fn pending_exception() -> Self {
        Self {
            kind: ErrorKind::PendingException,
            message: String::new(),
        }
    }
}

impl std::fmt::Display for EngineError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for EngineError {}

pub fn export_xlsx_structured_json(bytes: &[u8], options: &str) -> Result<String, EngineError> {
    session::export_xlsx_structured_json(bytes, options).map_err(EngineError::message)
}

pub fn export_xlsx_markdown_json(
    bytes: &[u8],
    options: &str,
    markdown_options: &str,
) -> Result<String, EngineError> {
    session::export_xlsx_markdown_json(bytes, options, markdown_options)
        .map_err(EngineError::message)
}

pub fn render_xlsx_markdown_json(content: &str, options: &str) -> Result<String, EngineError> {
    session::render_xlsx_markdown_json(content, options).map_err(EngineError::message)
}
