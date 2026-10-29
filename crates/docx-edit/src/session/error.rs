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
}

pub type EngineResult<T> = Result<T, EngineError>;

impl EngineError {
    pub fn message(error: impl fmt::Display) -> Self {
        Self {
            kind: EngineErrorKind::Message,
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
