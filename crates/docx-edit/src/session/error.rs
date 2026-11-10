use std::fmt;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EngineError {
    pub kind: EngineErrorKind,
    pub message: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EngineErrorKind {
    Message,
    PeerMetadata {
        code: &'static str,
    },
    SuggestedFormattingOverContainer,
    /// A table operation or save refused over a table that lost its payload to a concurrent undo.
    TableUnavailable,
}

pub type EngineResult<T> = Result<T, EngineError>;

impl EngineError {
    /// An error carrying `error`'s text. The text of an unavailable-table refusal keeps its kind,
    /// so every session path that refuses such a table reports it alike.
    pub fn message(error: impl fmt::Display) -> Self {
        let message = error.to_string();
        let kind = if message == crate::unavailable_tables::TABLE_UNAVAILABLE {
            EngineErrorKind::TableUnavailable
        } else {
            EngineErrorKind::Message
        };
        Self { kind, message }
    }
}

impl From<crate::OpError> for EngineError {
    fn from(error: crate::OpError) -> Self {
        let kind = match &error {
            crate::OpError::SuggestedFormattingOverContainer => {
                EngineErrorKind::SuggestedFormattingOverContainer
            }
            crate::OpError::TableUnavailable => EngineErrorKind::TableUnavailable,
            _ => EngineErrorKind::Message,
        };
        Self {
            kind,
            message: error.to_string(),
        }
    }
}

impl From<crate::PeerMetadataError> for EngineError {
    fn from(error: crate::PeerMetadataError) -> Self {
        let code = error.code();
        Self {
            kind: EngineErrorKind::PeerMetadata { code },
            message: error.to_string(),
        }
    }
}

impl fmt::Display for EngineError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.message)
    }
}

impl std::error::Error for EngineError {}
