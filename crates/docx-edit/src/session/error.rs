use std::fmt;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EngineError {
    pub kind: EngineErrorKind,
    pub message: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EngineErrorKind {
    Message,
    PeerMetadata { code: &'static str },
    SuggestedFormattingOverContainer,
    /// A write aimed at a table cell that a live table names while its story is missing or removed.
    CellHasNoStory,
}

pub type EngineResult<T> = Result<T, EngineError>;

impl EngineError {
    /// An error carrying `error`'s text. The text of a [`crate::EditError::CellHasNoStory`] or
    /// [`crate::OpError::CellHasNoStory`] keeps its kind, so every session path that refuses
    /// such a cell reports it alike.
    pub fn message(error: impl fmt::Display) -> Self {
        let message = error.to_string();
        let kind = if is_cell_has_no_story(&message) {
            EngineErrorKind::CellHasNoStory
        } else {
            EngineErrorKind::Message
        };
        Self { kind, message }
    }
}

/// Whether `message` is exactly the `table cell {id:?} has no story` text of a cell refusal.
fn is_cell_has_no_story(message: &str) -> bool {
    message
        .strip_prefix("table cell \"")
        .is_some_and(|rest| rest.ends_with("\" has no story"))
}

impl From<crate::OpError> for EngineError {
    fn from(error: crate::OpError) -> Self {
        let kind = match &error {
            crate::OpError::SuggestedFormattingOverContainer => {
                EngineErrorKind::SuggestedFormattingOverContainer
            }
            crate::OpError::CellHasNoStory(_) => EngineErrorKind::CellHasNoStory,
            _ => EngineErrorKind::Message,
        };
        Self {
            kind,
            message: error.to_string(),
        }
    }
}

impl From<crate::EditError> for EngineError {
    fn from(error: crate::EditError) -> Self {
        let kind = match &error {
            crate::EditError::CellHasNoStory(_) => EngineErrorKind::CellHasNoStory,
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
