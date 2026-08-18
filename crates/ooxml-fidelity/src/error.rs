use thiserror::Error;

#[derive(Debug, Error, PartialEq, Eq)]
pub enum FidelityError {
    #[error("malformed XML in {part}: {message}")]
    MalformedXml { part: String, message: String },
    #[error("unsafe XML in {part}: {kind}")]
    UnsafeXml { part: String, kind: &'static str },
    #[error("resource limit {kind} exceeded in {part}")]
    ResourceLimit { part: String, kind: &'static str },
    #[error("prefix {prefix:?} is not bound in {part}")]
    UnboundPrefix { part: String, prefix: String },
    #[error("artifact {0:?} has no registered comparison mode")]
    UnknownArtifact(String),
}
